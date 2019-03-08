let product =
	localStorage.getItem('theme-switcher-product') !== null ? localStorage.getItem('theme-switcher-product') : 'Flow';
localStorage.setItem('theme-switcher-product', product);
const themeNames = [
	'Darker',
	'Dark',
	'Dark Alt',
	'Primary',
	'Secondary',
	'Tertiary',
	'Light',
	'Lighter',
	'Lighter Alt'
];
const nodeTypes = ['fillNodes', 'strokeNodes', 'backgroundNodes'];
const propertyNames = ['inheritFillStyleID', 'inheritFillStyleIDForStroke', 'inheritFillStyleIDForBackground'];
let styleMapping;
let libraryStyles;

const convertSelection = () => {
	convert(figmaPlus.scene.selection);
};

const convertAll = () => {
	convert(figmaPlus.scene.currentPage);
};

const convert = async nodes => {
	if (!styleMapping) styleMapping = await getStyleMapping();
	const sortedNodes = sortNodes(nodes);
	changeThemeColor(sortedNodes);
};

const until = conditionFunction => {
	const poll = resolve => {
		if (conditionFunction()) resolve();
		else setTimeout(() => poll(resolve), 10);
	};

	return new Promise(poll);
};

const getStyleData = canvas_url => {
	return new Promise(resolve => {
		var xhr = new XMLHttpRequest();
		xhr.open('GET', canvas_url);
		xhr.responseType = 'arraybuffer';
		xhr.onload = () => {
			resolve(new Uint8Array(xhr.response));
		};
		xhr.send();
	});
};

const getStyleMapping = () => {
	const fileKey = 'H0LcCO1smJuohlD528oByhJb';
	const editingFileKey = App._state.editingFileKey;
	libraryStyles = Object.values(App._state.library.published.styles['588101129417313799'][fileKey]);
	libraryStyles = libraryStyles.filter(
		style =>
			style.name.startsWith('Grayscale/') ||
			style.name.startsWith('PowerApps Theme/') ||
			style.name.startsWith('Flow Theme/') ||
			style.name.startsWith('Power BI Theme/') ||
			style.name.startsWith('Other/')
	);
	const localStylePromises = libraryStyles.map(style => {
		return getStyleData(style.canvas_url).then(data => {
			const obj = {};
			obj[style.name] = App.sendMessage(
				'getOrCreateSubscribedStyleNodeId',
				{
					styleKey: style.key,
					fileKey: fileKey,
					editingFileKey: editingFileKey,
					versionHash: style.content_hash
				},
				data
			).args.localGUID;
			return obj;
		});
	});
	return new Promise(resolve => {
		Promise.all(localStylePromises).then(localStyleNodes => {
			resolve(
				localStyleNodes.reduce((style, x) => {
					for (var key in x)
						style[key] = { sessionID: parseInt(x[key].split(':')[0]), localID: parseInt(x[key].split(':')[1]) };
					return style;
				}, {})
			);
		});
	});
};

const nodeIdToObject = nodeId => {
	return { sessionID: parseInt(nodeId.split(':')[0]), localID: parseInt(nodeId.split(':')[1]) };
};

const getNodeStyle = node => {
	const nodeLog = DebuggingHelpers.logNode(node.id);
	const firstLine = nodeLog
		.match(/logging node state for \d+:\d+\n\n.+/)[0]
		.slice(nodeLog.match(/logging node state for \d+:\d+\n\n.+/)[0].indexOf('\n') + 2);
	if (firstLine.match(/\[.+\]/) !== null) {
		const ids = firstLine.match(/\d+:\d+/g);
		node.masterName = App._state.mirror.sceneGraph.get(ids[ids.length - 1]).name;
	}
	if (nodeLog.match(/inherit-fill-style-id: <GUID:\d+:\d+>/) !== null) {
		node.fillStyle = App._state.mirror.sceneGraph.get(
			nodeLog
				.match(/inherit-fill-style-id: <GUID:\d+:\d+>/)[0]
				.split(': ')[1]
				.match(/\d+:\d+/)[0]
		).name;
	}
	if (nodeLog.match(/inherit-fill-style-id-for-background: <GUID:\d+:\d+>/) !== null) {
		node.backgroundFillStyle = App._state.mirror.sceneGraph.get(
			nodeLog
				.match(/inherit-fill-style-id-for-background: <GUID:\d+:\d+>/)[0]
				.split(': ')[1]
				.match(/\d+:\d+/)[0]
		).name;
	}
	if (nodeLog.match(/inherit-fill-style-id-for-stroke: <GUID:\d+:\d+>/) !== null) {
		node.strokeFillStyle = App._state.mirror.sceneGraph.get(
			nodeLog
				.match(/inherit-fill-style-id-for-stroke: <GUID:\d+:\d+>/)[0]
				.split(': ')[1]
				.match(/\d+:\d+/)[0]
		).name;
	}
	if (nodeLog.match(/<FontHandle:.+>/) !== null) {
		node.isIcon = nodeLog.match(/<FontHandle:.+>/)[0].includes('MDL');
	}
	return node;
};

const sortNodes = nodes => {
	if (nodes.length !== undefined) {
		nodes = nodes
			.map(node => {
				let includedNodes = [node];
				if (node.children) includedNodes.push(...node.getAllDescendents());
				return includedNodes;
			})
			.reduce((a, b) => a.concat(b), []);
	} else nodes = nodes.getAllDescendents();
	nodes = nodes.map(getNodeStyle);

	const fillNodes = nodes.filter(
		node =>
			node.type !== 'TEXT' &&
			node.fillStyle &&
			((node.fillStyle.includes(' Theme/') && !node.fillStyle.includes(product)) ||
				node.fillStyle.startsWith('Grayscale/'))
	);

	const strokeNodes = nodes.filter(
		node =>
			node.type !== 'TEXT' &&
			node.strokeFillStyle &&
			((node.strokeFillStyle.includes(' Theme/') && !node.strokeFillStyle.includes(product)) ||
				node.strokeFillStyle.startsWith('Grayscale/'))
	);

	const backgroundNodes = nodes.filter(
		node =>
			node.type !== 'TEXT' &&
			node.backgroundFillStyle &&
			((node.backgroundFillStyle.includes(' Theme/') && !node.backgroundFillStyle.includes(product)) ||
				node.backgroundFillStyle.startsWith('Grayscale/'))
	);

	const whiteTextNodes = nodes.filter(
		node =>
			node.type === 'TEXT' &&
			node.fillStyle &&
			(node.fillStyle.includes('Grayscale/white') || node.fillStyle.includes('Grayscale/neutralLighterAlt'))
	);

	const blackTextNodes = nodes.filter(
		node =>
			node.type === 'TEXT' &&
			node.fillStyle &&
			(node.fillStyle.includes('Grayscale/neutralPrimary') ||
				node.fillStyle.includes('Grayscale/neutralDark') ||
				node.fillStyle.includes('Grayscale/black'))
	);

	const themeTextNodes = nodes.filter(
		node =>
			node.type === 'TEXT' && node.fillStyle && node.fillStyle.includes(' Theme/') && !node.fillStyle.includes(product)
	);

	let sortedNodes = {
		fillNodes: {},
		strokeNodes: {},
		backgroundNodes: {},
		textNodes: {
			white: whiteTextNodes,
			black: blackTextNodes,
			theme: {}
		}
	};

	themeNames.forEach(themeName => {
		Object.defineProperty(sortedNodes.fillNodes, themeName, {
			value: fillNodes.filter(node => node.fillStyle.endsWith(themeName))
		});
		Object.defineProperty(sortedNodes.strokeNodes, themeName, {
			value: strokeNodes.filter(node => node.strokeFillStyle.endsWith(themeName))
		});
		Object.defineProperty(sortedNodes.backgroundNodes, themeName, {
			value: backgroundNodes.filter(node => node.backgroundFillStyle.endsWith(themeName))
		});
		Object.defineProperty(sortedNodes.textNodes.theme, themeName, {
			value: themeTextNodes.filter(node => node.fillStyle.endsWith(themeName)),
			enumerable: true
		});
	});

	sortedNodes.fillNodes.white = fillNodes.filter(
		node => node.fillStyle.includes('Grayscale/white') || node.fillStyle.includes('Grayscale/neutralLighterAlt')
	);
	sortedNodes.fillNodes.black = fillNodes.filter(
		node =>
			node.fillStyle.includes('Grayscale/black') ||
			node.fillStyle.includes('Grayscale/neutralDark') ||
			node.fillStyle.includes('Grayscale/neutralPrimary')
	);

	sortedNodes.strokeNodes.white = strokeNodes.filter(
		node =>
			node.strokeFillStyle.includes('Grayscale/white') || node.strokeFillStyle.includes('Grayscale/neutralLighterAlt')
	);
	sortedNodes.strokeNodes.black = strokeNodes.filter(
		node =>
			node.strokeFillStyle.includes('Grayscale/black') ||
			node.strokeFillStyle.includes('Grayscale/neutralDark') ||
			node.strokeFillStyle.includes('Grayscale/neutralPrimary')
	);

	sortedNodes.backgroundNodes.white = backgroundNodes.filter(
		node =>
			node.backgroundFillStyle.includes('Grayscale/white') ||
			node.backgroundFillStyle.includes('Grayscale/neutralLighterAlt')
	);
	sortedNodes.backgroundNodes.black = backgroundNodes.filter(
		node =>
			node.backgroundFillStyle.includes('Grayscale/black') ||
			node.backgroundFillStyle.includes('Grayscale/neutralDark') ||
			node.backgroundFillStyle.includes('Grayscale/neutralPrimary')
	);

	return sortedNodes;
};

const changeThemeColor = async nodes => {
	figmaPlus.showToast(`Applying ${product} theme...`, 10);
	for (i = 0; i < nodeTypes.length; i++) {
		for (j = 0; j < themeNames.length; j++) {
			if (nodes[nodeTypes[i]][themeNames[j]].length > 0) {
				App.sendMessage('clearSelection');
				await until(() => App._state.mirror.selectionProperties.visible === null);
				App.sendMessage('addToSelection', { nodeIds: nodes[nodeTypes[i]][themeNames[j]].map(node => node.id) });
				await until(() => App._state.mirror.selectionProperties.visible !== null);
				let property = {};
				property[propertyNames[i]] = styleMapping[`${product} Theme/${themeNames[j]}`];
				App.updateSelectionProperties(property);
			}
		}
	}
	switch (product) {
		case 'Flow':
			for (i = 0; i < themeNames.length; i++) {
				if (nodes.textNodes.theme[themeNames[i]].length > 0) {
					App.sendMessage('clearSelection');
					await until(() => App._state.mirror.selectionProperties.visible === null);
					App.sendMessage('addToSelection', { nodeIds: nodes.textNodes.theme[themeNames[i]].map(node => node.id) });
					await until(() => App._state.mirror.selectionProperties.visible !== null);
					let property = {};
					property.inheritFillStyleID = styleMapping[`${product} Theme/${themeNames[i]}`];
					App.updateSelectionProperties(property);
				}
			}

			if (nodes.textNodes.theme.Secondary.length > 0) {
				App.sendMessage('clearSelection');
				await until(() => App._state.mirror.selectionProperties.visible === null);
				App.sendMessage('addToSelection', { nodeIds: nodes.textNodes.theme.Secondary.map(node => node.id) });
				await until(() => App._state.mirror.selectionProperties.visible !== null);
				let property = {};
				property.inheritFillStyleID = styleMapping['Flow Theme/Primary'];
				App.updateSelectionProperties(property);
			}
			break;
		case 'Power BI':
			if (nodes.fillNodes.Primary.length > 0) {
				const radioDotNodes = nodes.fillNodes.Primary.filter(
					node => node.masterName && node.masterName === 'Radio Dot'
				);
				if (radioDotNodes.length > 0) {
					App.sendMessage('clearSelection');
					await until(() => App._state.mirror.selectionProperties.visible === null);
					App.sendMessage('addToSelection', { nodeIds: radioDotNodes.map(node => node.id) });
					await until(() => App._state.mirror.selectionProperties.visible !== null);
					let property = {};
					property.inheritFillStyleID = styleMapping['Grayscale/neutralPrimary (323130)'];
					App.updateSelectionProperties(property);
				}
				const radioOutlineNodes = nodes.strokeNodes.Primary.filter(
					node => node.masterName && node.masterName === 'Radio Outline'
				);
				if (radioOutlineNodes.length > 0) {
					App.sendMessage('clearSelection');
					await until(() => App._state.mirror.selectionProperties.visible === null);
					App.sendMessage('addToSelection', { nodeIds: radioOutlineNodes.map(node => node.id) });
					await until(() => App._state.mirror.selectionProperties.visible !== null);
					let property = {};
					property.inheritFillStyleIDForStroke = styleMapping['Grayscale/neutralPrimary (323130)'];
					App.updateSelectionProperties(property);
				}
			}
			const focusBorderPrimary = nodes.strokeNodes.white.filter(
				node => node.masterName && node.masterName === 'Focus Border/Primary'
			);
			if (focusBorderPrimary.length > 0) {
				App.sendMessage('clearSelection');
				await until(() => App._state.mirror.selectionProperties.visible === null);
				App.sendMessage('addToSelection', { nodeIds: focusBorderPrimary.map(node => node.id) });
				await until(() => App._state.mirror.selectionProperties.visible !== null);
				let property = {};
				property.inheritFillStyleIDForStroke = styleMapping['Grayscale/neutralPrimary (323130)'];
				App.updateSelectionProperties(property);
			}

			const themeTextNodes = [].concat.apply([], Object.values(nodes.textNodes.theme));
			if (themeTextNodes.length > 0) {
				App.sendMessage('clearSelection');
				await until(() => App._state.mirror.selectionProperties.visible === null);
				App.sendMessage('addToSelection', { nodeIds: themeTextNodes.map(node => node.id) });
				await until(() => App._state.mirror.selectionProperties.visible !== null);
				let property = {};
				property.inheritFillStyleID = styleMapping['Grayscale/neutralPrimary (323130)'];
				App.updateSelectionProperties(property);
			}
			const linkNodes =
				nodes.textNodes.theme.Secondary.length > 0
					? nodes.textNodes.theme.Secondary.filter(node => node.masterName && node.masterName === 'Link')
					: [];
			if (linkNodes.length > 0) {
				App.sendMessage('clearSelection');
				await until(() => App._state.mirror.selectionProperties.visible === null);
				App.sendMessage('addToSelection', { nodeIds: linkNodes.map(node => node.id) });
				await until(() => App._state.mirror.selectionProperties.visible !== null);
				let property = {};
				property.inheritFillStyleID = styleMapping['Other/Link'];
				App.updateSelectionProperties(property);
			}
			if (nodes.textNodes.white.length > 0) {
				const buttonWhiteTextNodes = nodes.textNodes.white.filter(
					node =>
						node.masterName &&
						(node.masterName === 'Button Label' ||
							node.masterName === 'Button Description' ||
							node.masterName === 'Button Chevron' ||
							node.masterName === 'Button Icon' ||
							node.masterName === 'Checkmark')
				);
				if (buttonWhiteTextNodes.length > 0) {
					App.sendMessage('clearSelection');
					await until(() => App._state.mirror.selectionProperties.visible === null);
					App.sendMessage('addToSelection', { nodeIds: buttonWhiteTextNodes.map(node => node.id) });
					await until(() => App._state.mirror.selectionProperties.visible !== null);
					let property = {};
					property.inheritFillStyleID = styleMapping['Grayscale/neutralPrimary (323130)'];
					App.updateSelectionProperties(property);
				}
				const pressedButtonWhiteTextNodes = nodes.textNodes.white.filter(
					node => node.masterName && node.masterName === 'Button Label/Pressed'
				);
				if (pressedButtonWhiteTextNodes.length > 0) {
					App.sendMessage('clearSelection');
					await until(() => App._state.mirror.selectionProperties.visible === null);
					App.sendMessage('addToSelection', { nodeIds: pressedButtonWhiteTextNodes.map(node => node.id) });
					await until(() => App._state.mirror.selectionProperties.visible !== null);
					let property = {};
					property.inheritFillStyleID = styleMapping['Grayscale/black (000000)'];
					App.updateSelectionProperties(property);
				}
			}
			if (nodes.backgroundNodes.Primary.length > 0) {
				const headerBarNodes = nodes.backgroundNodes.Primary.filter(node => node.masterName === 'Header Bar');
				if (headerBarNodes.length > 0) {
					App.sendMessage('clearSelection');
					await until(() => App._state.mirror.selectionProperties.visible === null);
					App.sendMessage('addToSelection', { nodeIds: headerBarNodes.map(node => node.id) });
					await until(() => App._state.mirror.selectionProperties.visible !== null);
					let property = {};
					property.inheritFillStyleIDForBackground = styleMapping['Grayscale/neutralPrimaryAlt (3B3A39)'];
					App.updateSelectionProperties(property);
				}
			}
	}
	figmaPlus.showToast(`✔️ ${product} theme applied`);
};

figmaPlus.createContextMenuItem.Canvas(`Apply Flow theme to all`, convertAll, () => product === 'Flow');
figmaPlus.createContextMenuItem.Canvas(`Apply Power BI theme to all`, convertAll, () => product === 'Power BI');
figmaPlus.createContextMenuItem.Selection(`Apply Flow theme`, convertSelection, () => product === 'Flow');
figmaPlus.createContextMenuItem.Selection(`Apply Power BI theme`, convertSelection, () => product === 'Power BI');
figmaPlus.createPluginsMenuItem('Power Platform Theme Switcher', null, null, null, [
	{
		itemLabel: `Apply Flow theme to selection`,
		triggerFunction: convertSelection,
		condition: () => product === 'Flow' || figmaPlus.scene.selection.length < 1
	},
	{
		itemLabel: `Apply Power BI theme to selection`,
		triggerFunction: convertSelection,
		condition: () => product === 'Power BI' || figmaPlus.scene.selection.length < 1
	},
	{
		itemLabel: `Apply Flow theme to current page`,
		triggerFunction: convertAll,
		condition: () => product === 'Flow'
	},
	{
		itemLabel: `Apply Power BI theme to current page`,
		triggerFunction: convertAll,
		condition: () => product === 'Power BI'
	},
	{
		itemLabel: 'Set to use Flow theme',
		triggerFunction: () => {
			product = 'Flow';
			localStorage.setItem('theme-switcher-product', 'Flow');
			figmaPlus.showToast('Plugin settings have changed. Refresh this tab to start using Flow theme.');
		},
		condition: () => product === 'Power BI'
	},
	{
		itemLabel: 'Set to use Power BI',
		triggerFunction: () => {
			product = 'Power BI';
			localStorage.setItem('theme-switcher-product', 'Power BI');
			figmaPlus.showToast('Plugin settings have changed. Refresh this tab to start using Power BI theme.');
		},
		condition: () => product === 'Flow'
	}
]);

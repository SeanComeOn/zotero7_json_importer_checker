MakeItRed = {
	id: null,
	version: null,
	rootURI: null,
	initialized: false,
	addedElementIDs: [],
	lastExportPath: '', // 记忆上一次指定的导出文件夹路径

	init({ id, version, rootURI }) {
		if (this.initialized) return;
		this.id = id;
		this.version = version;
		this.rootURI = rootURI;
		this.initialized = true;
	},

	log(msg) {
		Zotero.debug("Make It Red: " + msg);
	},

	addToWindow(window) {
		let doc = window.document;
		let toolbar = this.findToolbarContainer(doc);
		if (!toolbar) {
			this.log('Toolbar container not found');
			return;
		}

		let style = doc.createElementNS('http://www.w3.org/1999/xhtml', 'style');
		style.id = 'wo-yao-yan-pai-json-toolbar-button-style';
		style.textContent = `
			#wo-yao-yan-pai-json-toolbar-button .toolbarbutton-icon,
			#wo-yao-yan-pai-export-toolbar-button .toolbarbutton-icon {
				width: 16px !important;
				height: 16px !important;
			}
			.wo-yao-yan-pai-html-btn {
				padding: 4px 10px;
				cursor: pointer;
				background-color: #f0f0f0;
				border: 1px solid #ccc;
				border-radius: 4px;
				font-family: inherit;
			}
			.wo-yao-yan-pai-html-btn:hover {
				background-color: #e0e0e0;
			}
		`;
		doc.documentElement.appendChild(style);
		this.storeAddedElement(style);

		// ==========================================
		// 1. 原有的 JSON 导入功能区
		// ==========================================
		let button = doc.createXULElement('toolbarbutton');
		button.id = 'wo-yao-yan-pai-json-toolbar-button';
		button.setAttribute('label', 'JSON Echo');
		button.setAttribute('tooltiptext', 'Open JSON echo panel');
		button.setAttribute('class', 'toolbarbutton-1');
		button.setAttribute('image', this.rootURI + 'icon.png');
		toolbar.appendChild(button);
		this.storeAddedElement(button);

		let panel = doc.createXULElement('panel');
		panel.id = 'wo-yao-yan-pai-json-panel';
		panel.setAttribute('type', 'arrow');
		panel.setAttribute('orient', 'vertical');
		panel.setAttribute('style', 'padding: 15px; width: 800px; max-height: 85vh; display: flex; flex-direction: column; box-sizing: border-box;');

		let vbox = doc.createXULElement('vbox');
		vbox.setAttribute('style', 'gap: 10px; flex: 1; display: flex; flex-direction: column; overflow: hidden;');
		vbox.setAttribute('flex', '1');

		let inputLabel = doc.createXULElement('label');
		inputLabel.setAttribute('value', 'Input JSON string (array of paper entries):');

		let input = doc.createElementNS('http://www.w3.org/1999/xhtml', 'textarea');
		input.id = 'wo-yao-yan-pai-json-input';
		input.setAttribute('rows', '5');
		input.setAttribute('style', 'width: 100%; box-sizing: border-box; font-family: monospace; resize: vertical; min-height: 80px;');

		let pathLabel = doc.createXULElement('label');
		pathLabel.setAttribute('value', 'Target collection path (example: My Library/Robot/Locomotion):');

		let collectionPathInput = doc.createElementNS('http://www.w3.org/1999/xhtml', 'input');
		collectionPathInput.id = 'wo-yao-yan-pai-collection-path';
		collectionPathInput.setAttribute('type', 'text');
		collectionPathInput.setAttribute('placeholder', 'Auto-filled from currently selected collection');
		collectionPathInput.setAttribute('style', 'width: 100%; box-sizing: border-box; padding: 4px;');

		let lockPathCheckbox = doc.createElementNS('http://www.w3.org/1999/xhtml', 'input');
		lockPathCheckbox.id = 'wo-yao-yan-pai-lock-collection-path';
		lockPathCheckbox.setAttribute('type', 'checkbox');
		lockPathCheckbox.title = 'Lock target collection path';

		let lockPathLabel = doc.createXULElement('label');
		lockPathLabel.setAttribute('value', 'Lock');
		lockPathLabel.setAttribute('control', 'wo-yao-yan-pai-lock-collection-path');

		let pathRow = doc.createXULElement('hbox');
		pathRow.setAttribute('style', 'gap: 8px; align-items: center;');
		collectionPathInput.setAttribute('flex', '1');
		pathRow.appendChild(collectionPathInput);
		pathRow.appendChild(lockPathCheckbox);
		pathRow.appendChild(lockPathLabel);

		let actions = doc.createXULElement('hbox');
		actions.setAttribute('style', 'gap: 8px;');

		let importButton = doc.createElementNS('http://www.w3.org/1999/xhtml', 'button');
		importButton.className = 'wo-yao-yan-pai-html-btn';
		importButton.textContent = 'Import by DOI';

		let clearButton = doc.createElementNS('http://www.w3.org/1999/xhtml', 'button');
		clearButton.className = 'wo-yao-yan-pai-html-btn';
		clearButton.textContent = 'Clear';

		let closeButton = doc.createElementNS('http://www.w3.org/1999/xhtml', 'button');
		closeButton.className = 'wo-yao-yan-pai-html-btn';
		closeButton.textContent = 'Close';

		actions.appendChild(importButton);
		actions.appendChild(clearButton);
		actions.appendChild(closeButton);

		let summaryLabel = doc.createXULElement('label');
		summaryLabel.id = 'wo-yao-yan-pai-summary';
		summaryLabel.setAttribute('value', 'Ready.');

		let outputLabel = doc.createXULElement('label');
		outputLabel.setAttribute('value', 'Validation list (Table view):');

		let output = doc.createElementNS('http://www.w3.org/1999/xhtml', 'div');
		output.id = 'wo-yao-yan-pai-json-output';
		output.setAttribute('style', 'width: 100%; flex: 1; min-height: 200px; overflow-y: auto; background-color: #f9f9f9; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;');

		importButton.addEventListener('click', async () => {
			importButton.disabled = true;
			clearButton.disabled = true;
			summaryLabel.setAttribute('value', 'Processing...');
			try {
				let targetCollection = this.resolveTargetCollection(window, collectionPathInput.value || '');
				let rows = await this.importFromJSONString(window, input.value ?? '', targetCollection);
				output.innerHTML = this.formatRowsToTable(rows);
				let successCount = rows.filter(row => row.success).length;
				summaryLabel.setAttribute('value', `Done: ${successCount}/${rows.length} succeeded.`);
			}
			catch (e) {
				output.innerHTML = `<div style="color: red; padding: 15px; font-weight: bold;">ERROR: ${this.escapeHTML(e.message || String(e))}</div>`;
				summaryLabel.setAttribute('value', 'Failed.');
				this.log(`Import failed: ${e}`);
			}
			finally {
				importButton.disabled = false;
				clearButton.disabled = false;
			}
		});

		clearButton.addEventListener('click', () => {
			input.value = '';
			output.innerHTML = '';
			summaryLabel.setAttribute('value', 'Ready.');
		});

		closeButton.addEventListener('click', () => {
			panel.hidePopup();
		});

		button.addEventListener('command', () => {
			if (panel.state === 'open' || panel.state === 'showing') {
				panel.hidePopup();
				return;
			}
			let selectedPath = this.getSelectedCollectionPath(window);
			if (!lockPathCheckbox.checked && selectedPath) {
				collectionPathInput.value = selectedPath;
			}
			panel.openPopup(button, 'after_start', 0, 6, false, false);
			window.setTimeout(() => input.focus(), 0);
		});

		vbox.appendChild(inputLabel);
		vbox.appendChild(input);
		vbox.appendChild(pathLabel);
		vbox.appendChild(pathRow);
		vbox.appendChild(actions);
		vbox.appendChild(summaryLabel);
		vbox.appendChild(outputLabel);
		vbox.appendChild(output);
		panel.appendChild(vbox);

		let popupSet = doc.getElementById('mainPopupSet') || doc.documentElement;
		popupSet.appendChild(panel);
		this.storeAddedElement(panel);

		if (typeof this.addExportPanelToWindow === 'function') {
			this.addExportPanelToWindow(window, toolbar, popupSet);
		} else {
			this.log('Export panel module not loaded');
		}
	},

	findToolbarContainer(doc) {
		return doc.getElementById('zotero-items-toolbar')
			|| doc.getElementById('zotero-toolbar')
			|| doc.getElementById('nav-bar');
	},

	normalizePath(text) {
		return String(text || '').trim().replace(/\\/g, '/').replace(/^\/+|\/+$/g, '');
	},

	getCollectionByIDCompat(collectionID, byID) {
		if (collectionID == null) return null;
		if (byID?.has(collectionID)) {
			return byID.get(collectionID);
		}

		if (typeof Zotero.Collections?.get === 'function') {
			try {
				let col = Zotero.Collections.get(collectionID);
				if (col) {
					if (byID) byID.set(col.id, col);
					return col;
				}
			}
			catch (e) {
				this.log(`Collections.get failed for ${collectionID}: ${e}`);
			}
		}

		return null;
	},

	getCollectionPathNames(collection, byID) {
		if (!collection) return [];

		let names = [];
		let current = collection;
		let guard = 0;
		while (current && guard < 100) {
			guard++;
			if (current.name) names.unshift(current.name);

			let parentID = current.parentID;
			if (!parentID) break;
			current = this.getCollectionByIDCompat(parentID, byID);
		}

		return names;
	},

	resolveTargetCollection(window, pathText) {
		let trimmed = (pathText || '').trim();
		if (!trimmed) {
			let selected = window.ZoteroPane.getSelectedCollection();
			if (!selected) {
				throw new Error('No target collection selected and no path provided');
			}
			return selected;
		}

		let normalizedPath = this.normalizePath(trimmed);
		let collections = this.getAllCollectionsCompat();
		if (!collections.length) {
			throw new Error('Unable to enumerate collections in this Zotero version. Please select a collection in the left pane and leave path empty.');
		}
		let byID = new Map();
		for (let col of collections) {
			byID.set(col.id, col);
		}

		let targetNorm = this.normalizePath(normalizedPath).toLowerCase();
		let targetParts = targetNorm.split('/').filter(Boolean);
		let candidateRows = [];

		for (let col of collections) {
			let names = this.getCollectionPathNames(col, byID);

			let relPath = this.normalizePath(names.join('/'));
			let relPathNorm = relPath.toLowerCase();
			let relParts = relPathNorm.split('/').filter(Boolean);
			let libName = Zotero.Libraries.getName(col.libraryID);
			let absPath = this.normalizePath(`${libName}/${relPath}`);
			let absPathNorm = absPath.toLowerCase();

			if (relPathNorm === targetNorm || absPathNorm === targetNorm) {
				return col;
			}

			candidateRows.push({ col, relParts, relPathNorm, absPathNorm, relPath });
		}

		// Fallback: allow matching by suffix when users provide partial hierarchical paths
		let candidateTailParts = [targetParts];
		if (targetParts.length > 1) {
			candidateTailParts.push(targetParts.slice(1));
		}

		let fallbackMatches = [];
		for (let row of candidateRows) {
			for (let tailParts of candidateTailParts) {
				if (!tailParts.length || row.relParts.length < tailParts.length) continue;
				let suffix = row.relParts.slice(row.relParts.length - tailParts.length).join('/');
				if (suffix === tailParts.join('/')) {
					fallbackMatches.push(row);
					break;
				}
			}
		}

		if (fallbackMatches.length === 1) {
			return fallbackMatches[0].col;
		}

		if (fallbackMatches.length > 1) {
			let preview = fallbackMatches.slice(0, 3).map(row => row.relPath).join(' | ');
			throw new Error(`Collection path is ambiguous: ${trimmed}. Matches: ${preview}`);
		}

		throw new Error(`Collection path not found: ${trimmed}`);
	},

	getSelectedCollectionPath(window) {
		let selected = window?.ZoteroPane?.getSelectedCollection?.();
		if (!selected) {
			return '';
		}

		let collections = this.getAllCollectionsCompat();
		let byID = new Map();
		for (let col of collections) {
			byID.set(col.id, col);
		}

		let names = this.getCollectionPathNames(selected, byID);

		let libName = Zotero.Libraries.getName(selected.libraryID) || '';
		return libName ? `${libName}/${names.join('/')}` : names.join('/');
	},

	getAllCollectionsCompat() {
		let all = [];
		let seen = new Set();
		let add = (collection) => {
			if (!collection || collection.id == null) return;
			if (seen.has(collection.id)) return;
			seen.add(collection.id);
			all.push(collection);
		};

		if (typeof Zotero.Collections?.getAll === 'function') {
			for (let col of Zotero.Collections.getAll() || []) {
				add(col);
			}
			if (all.length) return all;
		}

		if (typeof Zotero.Collections?.getByLibrary === 'function') {
			for (let lib of this.getAllLibrariesCompat()) {
				let libraryID = lib?.libraryID ?? lib?.id;
				if (libraryID == null) continue;
				let cols = [];
				try {
					cols = Zotero.Collections.getByLibrary(libraryID) || [];
				}
				catch (e) {
					this.log(`getByLibrary failed for library ${libraryID}: ${e}`);
				}
				for (let col of cols) {
					add(col);
				}

				// Some Zotero versions return only top-level collections here.
				let queue = [...cols];
				while (queue.length) {
					let current = queue.shift();
					if (!current) continue;

					let children = [];
					if (typeof current.getChildCollections === 'function') {
						try {
							children = current.getChildCollections() || [];
						}
						catch (e) {
							this.log(`getChildCollections failed for ${current.id}: ${e}`);
						}
					}

					for (let child of children) {
						if (child && !seen.has(child.id)) {
							add(child);
							queue.push(child);
						}
					}
				}
			}
		}

		return all;
	},

	getAllLibrariesCompat() {
		if (typeof Zotero.Libraries?.getAll === 'function') {
			return Zotero.Libraries.getAll() || [];
		}

		let libs = [];
		if (Zotero.Libraries?.userLibraryID != null) {
			let userLib = Zotero.Libraries.get(Zotero.Libraries.userLibraryID);
			if (userLib) libs.push(userLib);
		}
		return libs;
	},

	async importFromJSONString(window, jsonText, targetCollection) {
		let entries;
		try {
			entries = JSON.parse(jsonText);
		}
		catch {
			throw new Error('Invalid JSON');
		}

		if (!Array.isArray(entries)) {
			throw new Error('JSON must be an array of entries');
		}

		let rows = [];
		for (let [index, entry] of entries.entries()) {
			let doi = String(entry?.doi || entry?.DOI || '').trim();
			let jsonTitle = String(entry?.title || '').trim();

			if (!doi) {
				rows.push({ index, success: false, doi: '', foundTitle: '', jsonTitle, error: 'Missing DOI' });
				continue;
			}

			try {
				let item = await this.findOrCreateByDOI(doi, targetCollection);
				if (!item) {
					rows.push({ index, success: false, doi, foundTitle: '', jsonTitle, error: 'Identifier lookup failed or returned no item' });
					continue;
				}

				await this.addItemToCollection(targetCollection, item.id);
				rows.push({ index, success: true, doi, foundTitle: item.getField('title') || '', jsonTitle, error: '' });
			}
			catch (e) {
				rows.push({ index, success: false, doi, foundTitle: '', jsonTitle, error: String(e) });
			}
		}

		return rows;
	},

	async addItemToCollection(collection, itemID) {
		if (!collection || itemID == null) throw new Error('Invalid collection or item');

		let item = Zotero.Items.get(itemID);
		if (!item) return;

		let colIDs = item.getCollections();
		if (!colIDs.includes(collection.id)) {
			colIDs.push(collection.id);
			item.setCollections(colIDs);
			await item.saveTx();
		}
	},

	async findOrCreateByDOI(doi, targetCollection) {
		let libraryID = targetCollection.libraryID;
		let cleanDOI = Zotero.Utilities.cleanDOI(doi);
		if (!cleanDOI) {
			this.log(`Invalid DOI format: ${doi}`);
			return null;
		}

		let existing = await this.findExistingByDOI(cleanDOI, libraryID);
		if (existing) return existing;

		const fetchFromNetwork = async (identifier) => {
			let translate = new Zotero.Translate.Search();
			translate.setIdentifier(identifier);

			let translators = await translate.getTranslators();
			if (!translators || translators.length === 0) return null;

			translate.setTranslator(translators[0]);

			let newItems = [];
			translate.setHandler('itemDone', (obj, item) => newItems.push(item));

			try {
				let savedItems = await translate.translate({ libraryID: libraryID, collections: [targetCollection.id] });
				if (savedItems && savedItems.length > 0) return savedItems[0];
				if (newItems.length > 0 && newItems[0].id) return Zotero.Items.get(newItems[0].id);
				return null;
			}
			catch (e) {
				if (newItems.length > 0 && newItems[0].id) {
					return Zotero.Items.get(newItems[0].id);
				}
				return null;
			}
		};

		let item = await fetchFromNetwork({ DOI: cleanDOI });
		if (item) return item;

		let arxivMatch = cleanDOI.match(/arxiv\.(.+)$/i);
		if (arxivMatch) {
			item = await fetchFromNetwork(`arXiv:${arxivMatch[1]}`);
			if (item) {
				if (!item.getField('DOI')) {
					item.setField('DOI', cleanDOI);
					await item.saveTx();
				}
				return item;
			}
		}

		let fallbackItem = await fetchFromNetwork(`doi:${cleanDOI}`);
		if (fallbackItem) return fallbackItem;

		return null;
	},

	async findExistingByDOI(doi, libraryID) {
		let search = new Zotero.Search();
		search.libraryID = libraryID;
		search.addCondition('DOI', 'is', doi);
		let itemIDs = await search.search();
		if (!itemIDs || !itemIDs.length) return null;
		return Zotero.Items.get(itemIDs[0]);
	},

	escapeHTML(value) {
		return String(value ?? '')
			.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;')
			.replace(/'/g, '&#39;');
	},

	formatRowsToTable(rows) {
		let html = `
		<table style="width: 100%; border-collapse: collapse; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; font-size: 13px;">
			<thead style="position: sticky; top: 0; background-color: #eaeaea; box-shadow: 0 1px 2px rgba(0,0,0,0.1); z-index: 1;">
				<tr>
					<th style="padding: 8px; border: 1px solid #ccc; text-align: left; width: 10%;">状态</th>
					<th style="padding: 8px; border: 1px solid #ccc; text-align: left; width: 15%;">DOI</th>
					<th style="padding: 8px; border: 1px solid #ccc; text-align: left; width: 30%;">实际抓取标题 (Zotero)</th>
					<th style="padding: 8px; border: 1px solid #ccc; text-align: left; width: 30%;">输入标题 (JSON)</th>
					<th style="padding: 8px; border: 1px solid #ccc; text-align: center; width: 15%;">对比结果</th>
				</tr>
			</thead>
			<tbody>
		`;

		for (let row of rows) {
			let statusColor = row.success ? '#2e7d32' : '#d32f2f';
			let statusText = row.success ? '成功' : '失败';
			let reason = row.error ? `<br/><span style="color: #d32f2f; font-size: 11px; font-weight: normal;">${this.escapeHTML(row.error)}</span>` : '';

			let matchResult = '';
			if (row.success) {
				let cleanFound = (row.foundTitle || '').toLowerCase().replace(/[^a-z0-9]/g, '');
				let cleanJson = (row.jsonTitle || '').toLowerCase().replace(/[^a-z0-9]/g, '');
				if (cleanFound !== '' && cleanFound === cleanJson) {
					matchResult = '<span style="color: #2e7d32; font-weight: bold; background: #e8f5e9; padding: 3px 6px; border-radius: 4px;">标题匹配</span>';
				} else {
					matchResult = '<span style="color: #d32f2f; font-weight: bold; background: #ffebee; padding: 3px 6px; border-radius: 4px;">不匹配<br/><span style="font-size: 11px; font-weight: normal;">(疑似幻觉)</span></span>';
				}
			} else {
				matchResult = '<span style="color: #999;">-</span>';
			}

			html += `
				<tr style="background-color: #fff; border-bottom: 1px solid #eee;">
					<td style="padding: 8px; border: 1px solid #eee; color: ${statusColor}; font-weight: bold; vertical-align: top;">${statusText}${reason}</td>
					<td style="padding: 8px; border: 1px solid #eee; word-break: break-all; vertical-align: top;">${this.escapeHTML(row.doi || '-')}</td>
					<td style="padding: 8px; border: 1px solid #eee; vertical-align: top; color: #333;">${this.escapeHTML(row.foundTitle || '-')}</td>
					<td style="padding: 8px; border: 1px solid #eee; vertical-align: top; color: #666;">${this.escapeHTML(row.jsonTitle || '-')}</td>
					<td style="padding: 8px; border: 1px solid #eee; text-align: center; vertical-align: middle;">${matchResult}</td>
				</tr>
			`;
		}

		html += `</tbody></table>`;
		return html;
	},

	addToAllWindows() {
		var windows = Zotero.getMainWindows();
		for (let win of windows) {
			if (!win.ZoteroPane) continue;
			this.addToWindow(win);
		}
	},

	storeAddedElement(elem) {
		if (!elem.id) throw new Error("Element must have an id");
		this.addedElementIDs.push(elem.id);
	},

	removeFromWindow(window) {
		var doc = window.document;
		for (let id of this.addedElementIDs) {
			doc.getElementById(id)?.remove();
		}
	},

	removeFromAllWindows() {
		var windows = Zotero.getMainWindows();
		for (let win of windows) {
			if (!win.ZoteroPane) continue;
			this.removeFromWindow(win);
		}
	},

	async main() {
		var host = new URL('https://foo.com/path').host;
		this.log(`Host is ${host}`);
	},
};
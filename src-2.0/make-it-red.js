MakeItRed = {
	id: null,
	version: null,
	rootURI: null,
	initialized: false,
	addedElementIDs: [],

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

		// 动态注入一段 CSS，强制限制我们按钮内部的图标尺寸
		let style = doc.createElementNS('http://www.w3.org/1999/xhtml', 'style');
		style.id = 'make-it-red-json-toolbar-button-style';
		style.textContent = `
			#make-it-red-json-toolbar-button .toolbarbutton-icon {
				width: 16px !important;
				height: 16px !important;
			}
		`;
		doc.documentElement.appendChild(style);
		this.storeAddedElement(style); // 确保插件卸载时能自动清理这段 CSS

		let button = doc.createXULElement('toolbarbutton');
		button.id = 'make-it-red-json-toolbar-button';
		button.setAttribute('label', 'JSON Echo');
		button.setAttribute('tooltiptext', 'Open JSON echo panel');
		button.setAttribute('class', 'toolbarbutton-1');
		button.setAttribute('image', this.rootURI + 'icon.png');
		toolbar.appendChild(button);
		this.storeAddedElement(button);

		let panel = doc.createXULElement('panel');
		panel.id = 'make-it-red-json-panel';
		panel.setAttribute('type', 'arrow');
		panel.setAttribute('orient', 'vertical');
		panel.setAttribute('style', 'padding: 15px; width: 800px; max-height: 85vh; display: flex; flex-direction: column; box-sizing: border-box;');

		let vbox = doc.createXULElement('vbox');
		vbox.setAttribute('style', 'gap: 10px; flex: 1; display: flex; flex-direction: column; overflow: hidden;');
		vbox.setAttribute('flex', '1');

		let inputLabel = doc.createXULElement('label');
		inputLabel.setAttribute('value', 'Input JSON string (array of paper entries):');

		let input = doc.createElementNS('http://www.w3.org/1999/xhtml', 'textarea');
		input.id = 'make-it-red-json-input';
		input.setAttribute('rows', '5');
		input.setAttribute('style', 'width: 100%; box-sizing: border-box; font-family: monospace; resize: vertical; min-height: 80px;');

		let pathLabel = doc.createXULElement('label');
		pathLabel.setAttribute('value', 'Target collection path (example: My Library/Robot/Locomotion):');

		let collectionPathInput = doc.createElementNS('http://www.w3.org/1999/xhtml', 'input');
		collectionPathInput.id = 'make-it-red-collection-path';
		collectionPathInput.setAttribute('type', 'text');
		collectionPathInput.setAttribute('placeholder', 'Auto-filled from currently selected collection');
		collectionPathInput.setAttribute('style', 'width: 100%; box-sizing: border-box; padding: 4px;');

		let lockPathCheckbox = doc.createElementNS('http://www.w3.org/1999/xhtml', 'input');
		lockPathCheckbox.id = 'make-it-red-lock-collection-path';
		lockPathCheckbox.setAttribute('type', 'checkbox');
		lockPathCheckbox.title = 'Lock target collection path';

		let lockPathLabel = doc.createXULElement('label');
		lockPathLabel.setAttribute('value', 'Lock');
		lockPathLabel.setAttribute('control', 'make-it-red-lock-collection-path');

		let pathRow = doc.createXULElement('hbox');
		pathRow.setAttribute('style', 'gap: 8px; align-items: center;');
		collectionPathInput.setAttribute('flex', '1');
		pathRow.appendChild(collectionPathInput);
		pathRow.appendChild(lockPathCheckbox);
		pathRow.appendChild(lockPathLabel);

		let actions = doc.createXULElement('hbox');
		actions.setAttribute('style', 'gap: 8px;');

		let importButton = doc.createXULElement('button');
		importButton.setAttribute('label', 'Import by DOI');

		let clearButton = doc.createXULElement('button');
		clearButton.setAttribute('label', 'Clear');

		let closeButton = doc.createXULElement('button');
		closeButton.setAttribute('label', 'Close');

		actions.appendChild(importButton);
		actions.appendChild(clearButton);
		actions.appendChild(closeButton);

		let summaryLabel = doc.createXULElement('label');
		summaryLabel.id = 'make-it-red-summary';
		summaryLabel.setAttribute('value', 'Ready.');

		let outputLabel = doc.createXULElement('label');
		outputLabel.setAttribute('value', 'Validation list (Table view):');

		let output = doc.createElementNS('http://www.w3.org/1999/xhtml', 'div');
		output.id = 'make-it-red-json-output';
		output.setAttribute('style', 'width: 100%; flex: 1; min-height: 200px; overflow-y: auto; background-color: #f9f9f9; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;');

		importButton.addEventListener('command', async () => {
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

		clearButton.addEventListener('command', () => {
			input.value = '';
			output.innerHTML = '';
			summaryLabel.setAttribute('value', 'Ready.');
		});

		closeButton.addEventListener('command', () => {
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

		input.addEventListener('keydown', (event) => {
			if (event.key === 'Escape') {
				panel.hidePopup();
				return;
			}
			event.stopPropagation();
		});

		collectionPathInput.addEventListener('keydown', (event) => {
			if (event.key === 'Escape') {
				panel.hidePopup();
				return;
			}
			event.stopPropagation();
		});

		collectionPathInput.addEventListener('click', () => {
			collectionPathInput.focus();
		});

		lockPathCheckbox.addEventListener('change', () => {
			if (lockPathCheckbox.checked) {
				return;
			}
			let selectedPath = this.getSelectedCollectionPath(window);
			if (selectedPath) {
				collectionPathInput.value = selectedPath;
			}
		});

		output.addEventListener('keydown', (event) => {
			if (event.key === 'Escape') {
				panel.hidePopup();
			}
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
	},

	findToolbarContainer(doc) {
		return doc.getElementById('zotero-items-toolbar')
			|| doc.getElementById('zotero-toolbar')
			|| doc.getElementById('nav-bar');
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

		let normalizedPath = trimmed.replace(/\\/g, '/').replace(/^\/+|\/+$/g, '');
		let collections = this.getAllCollectionsCompat();
		if (!collections.length) {
			throw new Error('Unable to enumerate collections in this Zotero version. Please select a collection in the left pane and leave path empty.');
		}
		let byID = new Map();
		for (let col of collections) {
			byID.set(col.id, col);
		}

		let normalize = (s) => (s || '').trim().replace(/\\/g, '/').replace(/^\/+|\/+$/g, '');
		let targetNorm = normalize(normalizedPath).toLowerCase();

		for (let col of collections) {
			let names = [col.name];
			let parentID = col.parentID;
			while (parentID) {
				let parent = byID.get(parentID);
				if (!parent) break;
				names.unshift(parent.name);
				parentID = parent.parentID;
			}

			let relPath = normalize(names.join('/'));
			let libName = Zotero.Libraries.getName(col.libraryID);
			let absPath = normalize(`${libName}/${relPath}`);

			if (relPath.toLowerCase() === targetNorm || absPath.toLowerCase() === targetNorm) {
				return col;
			}
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

		let names = [selected.name];
		let parentID = selected.parentID;
		while (parentID) {
			let parent = byID.get(parentID);
			if (!parent) break;
			names.unshift(parent.name);
			parentID = parent.parentID;
		}

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
				rows.push({
					index,
					success: false,
					doi: '',
					foundTitle: '',
					jsonTitle,
					error: 'Missing DOI'
				});
				continue;
			}

			try {
				let item = await this.findOrCreateByDOI(doi, targetCollection);
				if (!item) {
					rows.push({
						index,
						success: false,
						doi,
						foundTitle: '',
						jsonTitle,
						error: 'Identifier lookup failed or returned no item'
					});
					continue;
				}

				await this.addItemToCollection(targetCollection, item.id);
				rows.push({
					index,
					success: true,
					doi,
					foundTitle: item.getField('title') || '',
					jsonTitle,
					error: ''
				});
			}
			catch (e) {
				rows.push({
					index,
					success: false,
					doi,
					foundTitle: '',
					jsonTitle,
					error: String(e)
				});
			}
		}

		return rows;
	},

	async addItemToCollection(collection, itemID) {
		if (!collection || itemID == null) {
			throw new Error('Invalid collection or item');
		}

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
		if (existing) {
			return existing;
		}

		// 改进版的内部抓取函数
		const fetchFromNetwork = async (identifier) => {
			let translate = new Zotero.Translate.Search();
			translate.setIdentifier(identifier);

			let translators = await translate.getTranslators();
			if (!translators || translators.length === 0) return null;

			translate.setTranslator(translators[0]);

			let newItems = [];
			translate.setHandler('itemDone', (obj, item) => {
				newItems.push(item);
			});

			try {
				// 正常情况下，translate() 返回的数组里装的是真正的 Zotero.Item 对象
				let savedItems = await translate.translate({
					libraryID: libraryID,
					collections: [targetCollection.id]
				});

				if (savedItems && savedItems.length > 0) {
					return savedItems[0];
				}

				// 极端情况防御：如果 savedItems 为空但 itemDone 触发了
				if (newItems.length > 0 && newItems[0].id) {
					return Zotero.Items.get(newItems[0].id);
				}

				return null;
			}
			catch (e) {
				this.log(`Network lookup warning/error for ${JSON.stringify(identifier)}: ${e}`);

				// 🌟 终极修复：通过数据库 ID 获取真正的 Zotero.Item 实例
				if (newItems.length > 0 && newItems[0].id) {
					this.log(`Recovered item ${newItems[0].id} despite translation error (likely PDF failure).`);
					// Zotero.Items.get 返回的对象才真正拥有 getField 和 setField 方法！
					return Zotero.Items.get(newItems[0].id);
				}

				return null;
			}
		};

		// 1. Standard DOI search
		let item = await fetchFromNetwork({ DOI: cleanDOI });
		if (item) return item;

		// 2. arXiv fallback
		let arxivMatch = cleanDOI.match(/arxiv\.(.+)$/i);
		if (arxivMatch) {
			let arxivId = arxivMatch[1];
			item = await fetchFromNetwork(`arXiv:${arxivId}`);
			if (item) {
				if (!item.getField('DOI')) {
					item.setField('DOI', cleanDOI);
					await item.saveTx();
				}
				return item;
			}
		}

		// 3. Final fallback
		let fallbackItem = await fetchFromNetwork(`doi:${cleanDOI}`);
		if (fallbackItem) return fallbackItem;

		return null;
	},

	async findExistingByDOI(doi, libraryID) {
		let search = new Zotero.Search();
		search.libraryID = libraryID;
		search.addCondition('DOI', 'is', doi);
		let itemIDs = await search.search();
		if (!itemIDs || !itemIDs.length) {
			return null;
		}
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
			let reason = row.error
				? `<br/><span style="color: #d32f2f; font-size: 11px; font-weight: normal;">${this.escapeHTML(row.error)}</span>`
				: '';

			let matchResult = '';
			if (row.success) {
				let cleanFound = (row.foundTitle || '').toLowerCase().replace(/[^a-z0-9]/g, '');
				let cleanJson = (row.jsonTitle || '').toLowerCase().replace(/[^a-z0-9]/g, '');

				if (cleanFound !== '' && cleanFound === cleanJson) {
					matchResult = '<span style="color: #2e7d32; font-weight: bold; background: #e8f5e9; padding: 3px 6px; border-radius: 4px;">标题匹配</span>';
				}
				else {
					matchResult = '<span style="color: #d32f2f; font-weight: bold; background: #ffebee; padding: 3px 6px; border-radius: 4px;">不匹配<br/><span style="font-size: 11px; font-weight: normal;">(疑似幻觉)</span></span>';
				}
			}
			else {
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
		if (!elem.id) {
			throw new Error("Element must have an id");
		}
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
		// Global properties are included automatically in Zotero 7
		var host = new URL('https://foo.com/path').host;
		this.log(`Host is ${host}`);
	},
};

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

		let button = doc.createXULElement('toolbarbutton');
		button.id = 'make-it-red-json-toolbar-button';
		button.setAttribute('label', 'JSON Echo');
		button.setAttribute('tooltiptext', 'Open JSON echo panel');
		button.setAttribute('class', 'toolbarbutton-1');
		toolbar.appendChild(button);
		this.storeAddedElement(button);

		let panel = doc.createXULElement('panel');
		panel.id = 'make-it-red-json-panel';
		panel.setAttribute('type', 'arrow');
		panel.setAttribute('orient', 'vertical');
		panel.setAttribute('style', 'padding: 10px; width: 760px; max-height: 620px;');

		let vbox = doc.createXULElement('vbox');
		vbox.setAttribute('style', 'gap: 8px;');
		vbox.setAttribute('flex', '1');

		let inputLabel = doc.createXULElement('label');
		inputLabel.setAttribute('value', 'Input JSON string (array of paper entries):');

		let input = doc.createElementNS('http://www.w3.org/1999/xhtml', 'textarea');
		input.id = 'make-it-red-json-input';
		input.setAttribute('rows', '8');
		input.setAttribute('style', 'width: 740px; font-family: monospace;');

		let pathLabel = doc.createXULElement('label');
		pathLabel.setAttribute('value', 'Target collection path (example: My Library/Robot/Locomotion):');

		let collectionPathInput = doc.createElementNS('http://www.w3.org/1999/xhtml', 'input');
		collectionPathInput.id = 'make-it-red-collection-path';
		collectionPathInput.setAttribute('type', 'text');
		collectionPathInput.setAttribute('placeholder', 'Leave empty to use currently selected collection');
		collectionPathInput.setAttribute('style', 'width: 740px; box-sizing: border-box;');

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
		outputLabel.setAttribute('value', 'Validation list (status | found title | JSON title | DOI):');

		let output = doc.createElementNS('http://www.w3.org/1999/xhtml', 'textarea');
		output.id = 'make-it-red-json-output';
		output.setAttribute('rows', '10');
		output.setAttribute('readonly', 'readonly');
		output.setAttribute('style', 'width: 740px; font-family: monospace;');

		importButton.addEventListener('command', async () => {
			importButton.disabled = true;
			clearButton.disabled = true;
			summaryLabel.setAttribute('value', 'Processing...');
			try {
				let targetCollection = this.resolveTargetCollection(window, collectionPathInput.value || '');
				let rows = await this.importFromJSONString(window, input.value ?? '', targetCollection);
				output.value = this.formatRows(rows);
				let successCount = rows.filter(row => row.success).length;
				summaryLabel.setAttribute('value', `Done: ${successCount}/${rows.length} succeeded.`);
			}
			catch (e) {
				output.value = `ERROR: ${e.message}`;
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
			output.value = '';
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

		output.addEventListener('keydown', (event) => {
			if (event.key === 'Escape') {
				panel.hidePopup();
			}
		});

		vbox.appendChild(inputLabel);
		vbox.appendChild(input);
		vbox.appendChild(pathLabel);
		vbox.appendChild(collectionPathInput);
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
				let item = await this.findOrCreateByDOI(doi, targetCollection.libraryID);
				if (!item) {
					rows.push({
						index,
						success: false,
						doi,
						foundTitle: '',
						jsonTitle,
						error: 'Identifier lookup failed'
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

		if (typeof collection.hasItem === 'function' && collection.hasItem(itemID)) {
			return;
		}

		if (typeof Zotero.DB?.executeTransaction === 'function') {
			await Zotero.DB.executeTransaction(async () => {
				await collection.addItem(itemID);
			});
			return;
		}

		await collection.addItem(itemID);
	},

	async findOrCreateByDOI(doi, libraryID) {
		let existing = await this.findExistingByDOI(doi, libraryID);
		if (existing) return existing;

		let forms = [
			doi,
			`doi:${doi}`,
			{ DOI: doi }
		];

		for (let form of forms) {
			let translate = new Zotero.Translate.Search();
			try {
				translate.setIdentifier(form);
				let translators = await translate.getTranslators();
				if (!translators || !translators.length) {
					continue;
				}

				translate.setTranslator(translators);
				await translate.translate({ libraryID });
			}
			catch (e) {
				this.log(`Translate failed for DOI ${doi} with form ${JSON.stringify(form)}: ${e}`);
			}

			existing = await this.findExistingByDOI(doi, libraryID);
			if (existing) {
				return existing;
			}
		}

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

	formatRows(rows) {
		return rows.map((row) => {
			let status = row.success ? 'OK' : 'FAIL';
			let reason = row.error ? ` | reason: ${row.error}` : '';
			return `${status} | found: ${row.foundTitle || '(none)'} | json: ${row.jsonTitle || '(none)'} | doi: ${row.doi || '(none)'}${reason}`;
		}).join('\n');
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

		// Retrieve a global pref
		this.log(`Intensity is ${Zotero.Prefs.get('extensions.make-it-red.intensity', true)}`);
	},
};

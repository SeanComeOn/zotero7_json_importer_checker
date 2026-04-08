var LlmJsonImporter;

function log(msg) {
	Zotero.debug("LLM JSON Importer: " + msg);
}

function install() {
	log("Installed 1.0.0");
}

async function startup({ id, version, rootURI }) {
	log("Starting 1.0.0");

	// 注意这里加载的是重命名后的新文件
	Services.scriptloader.loadSubScript(rootURI + 'llm-json-importer.js');
	LlmJsonImporter.init({ id, version, rootURI });
	LlmJsonImporter.addToAllWindows();
	await LlmJsonImporter.main();
}

function onMainWindowLoad({ window }) {
	LlmJsonImporter.addToWindow(window);
}

function onMainWindowUnload({ window }) {
	LlmJsonImporter.removeFromWindow(window);
}

function shutdown() {
	log("Shutting down");
	LlmJsonImporter.removeFromAllWindows();
	LlmJsonImporter = undefined;
}

function uninstall() {
	log("Uninstalled");
}

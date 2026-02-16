const { app, core } = require("photoshop");
const fs = require("fs");
const uxp = require("uxp");

// Пути к папкам со шрифтами
const FONT_PATHS = {
	system: "file:C:/Windows/Fonts",
	adobe: "file:C:/Program Files/Common Files/Adobe/Fonts"
};

const tabNames = new Map([
	["all", "Все"],
	["system", "Системные"],
	["adobe", "Adobe"],
	["custom", "Другие"],
	["other", "Другие"]
]);

// Кэш шрифтов
let fontCache = {
	system: new Map(),
	adobe: new Map(),
	custom: new Map(),
	all: new Map()
};

function getExtension(filename) {
	return filename.slice(
		((filename.lastIndexOf('.') - 1) >>> 0) + 2
	);
}

function hasExtension(filename) {
	return getExtension(filename) !== '';
}

// Получить список файлов из папки

// Получить список шрифтов из папки
async function getFontFiles(path) {
	try {
		const entries = await fs.readdir(path);
		return entries.filter(file =>
			[".ttf", ".otf"].some(ext => file.toLowerCase().endsWith(ext))
		);
	} catch (error) {
		console.error(`Ошибка чтения шрифтов ${path}:`, error);
		return [];
	}
}

async function getFolders(path) {
	try {
		const entries = await fs.readdir(path);
		return entries.filter(file => !hasExtension(file));
	} catch (error) {
		console.error(`Ошибка чтения папок ${path}:`, error);
		return [];
	}
}

function getFileName(filename) {
	return filename
		.replace(/\.[^/.]+$/, "") // Удаляем расширение
		.trim();
}

// Получить чистое название шрифта из имени файла
function getCleanFontName(filename) {
	return getFileName(filename)
		.replace(/[-_]/g, " ")    // Заменяем разделители на пробелы
		.replace(/\d+/g, "")      // Удаляем цифры
		.trim();
}

async function getFontNameFromFile(file) {
	var text = await fs.readFile(file);
	const dataView = new DataView(text);
	return parseFontName(dataView);
}

function int32ToHexText(int32) {
	// Преобразуем в беззнаковое 32-битное число
	int32 = int32 >>> 0;

	// Извлекаем 4 байта
	const byte1 = (int32 >>> 24) & 0xFF;
	const byte2 = (int32 >>> 16) & 0xFF;
	const byte3 = (int32 >>> 8) & 0xFF;
	const byte4 = int32 & 0xFF;

	// Форматируем байты в HEX (двузначные, с ведущим нулём)
	const hex1 = byte1.toString(16).padStart(2, '0').toUpperCase();
	const hex2 = byte2.toString(16).padStart(2, '0').toUpperCase();
	const hex3 = byte3.toString(16).padStart(2, '0').toUpperCase();
	const hex4 = byte4.toString(16).padStart(2, '0').toUpperCase();
	return hex1 + " " + hex2 + " " + hex3 + " " + hex4;
}

function int32ToAsciiWithHex(int32) {
	// Преобразуем в беззнаковое 32-битное число
	int32 = int32 >>> 0;

	// Извлекаем 4 байта
	const byte1 = (int32 >>> 24) & 0xFF;
	const byte2 = (int32 >>> 16) & 0xFF;
	const byte3 = (int32 >>> 8) & 0xFF;
	const byte4 = int32 & 0xFF;

	// Преобразуем байты в символы ASCII
	const char1 = String.fromCharCode(byte1);
	const char2 = String.fromCharCode(byte2);
	const char3 = String.fromCharCode(byte3);
	const char4 = String.fromCharCode(byte4);
	const text = char1 + char2 + char3 + char4;

	return text + " (" + int32ToHexText(int32) + ")";
}

function parseFontName(dataView) {
	// Проверка, что это TTF (0x00010000) или OTF ('OTTO')
	const sfntVersion = dataView.getUint32(0, false);
	if (sfntVersion !== 0x00010000 && sfntVersion !== 0x4F54544F) {
		showAlert("Not a valid TTF/OTF file");
		throw new Error("Not a valid TTF/OTF file");
	}

	// Поиск таблицы 'name' (тег 0x6E616D65)
	const numTables = dataView.getUint16(4, false);
	let nameTableOffset = 0;
	let nameTableLength = 0;

	for (let i = 0; i < numTables; i++) {
		const offset = 12 + i * 16;
		const tag = dataView.getUint32(offset, false);
		if (tag === 0x6E616D65) { // 'name'
			nameTableOffset = dataView.getUint32(offset + 8, false);
			nameTableLength = dataView.getUint32(offset + 12, false);
			break;
		}
	}

	if (!nameTableOffset) {
		showAlert("No 'name' table found");
		throw new Error("No 'name' table found");
	}

	// Чтение таблицы 'name'
	const nameTable = new DataView(
		dataView.buffer,
		nameTableOffset,
		nameTableLength
	);

	const count = nameTable.getUint16(2, false);
	const storageOffset = nameTable.getUint16(4, false);

	// Поиск названия (nameID = 1 - Family, 4 - Full Name)
	for (let i = 0; i < count; i++) {
		const recordOffset = 6 + i * 12;
		const nameID = nameTable.getUint16(recordOffset + 6, false);

		if (nameID === 1 || nameID === 4) {
			const platformID = nameTable.getUint16(recordOffset, false);

			if (platformID === 3 || platformID === 1) {
				const length = nameTable.getUint16(recordOffset + 8, false);
				const offset = nameTable.getUint16(recordOffset + 10, false);

				const start = nameTableOffset + storageOffset + offset;
				let fontName = "";

				// Декодировка UTF-16BE (Windows) или MacRoman (упрощённо)
				switch (platformID) {
					case 1: // Mac (ASCII)
						for (let j = 0; j < length; j++) {
							fontName += String.fromCharCode(dataView.getUint8(start + j));
						}
						break;

					case 3: // Windows (Unicode)
						for (let j = 0; j < length; j += 2) {
							fontName += String.fromCharCode(dataView.getUint16(start + j, false));
						}
						break;
				}

				if (fontName) return fontName;
			}
		}
	}

	showAlert("Font name not found");
	throw new Error("Font name not found");
}

// Получить чистое название шрифта из имени файла
function getFirstWord(filename) {
	console.warn("filename:", filename);
	return getCleanFontName(filename).split(' ', 1).find(word => word.length > 1);
}

function deepNormalizeString(str) {
	return str
		// Нормализация Unicode (попробуйте разные формы)
		.normalize('NFKC') // NFKC более агрессивная, чем NFC
		// Удаление ВСЕХ невидимых и управляющих символов
		.replace(/[\u0000-\u001F\u007F-\u009F\u200B-\u200F\uFEFF\uFFFE\uFFFF]/g, '')
		// Замена всех типов пробелов на обычные
		.replace(/[\s\u00A0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000]/g, ' ')
		// Удаление лишних пробелов
		.replace(/\s+/g, ' ')
		.trim()
		// Приведение к нижнему регистру для надежности
		.toLowerCase();
}

// Сравнить шрифты Photoshop с файлами в папках
async function analyzeFonts() {
	var adobeFolders = await getFolders(FONT_PATHS.adobe);
	adobeFolders.unshift("");

	const psFonts = await app.fonts;

	const fontMap = new Map();

	for (const font of psFonts) {
		const key = deepNormalizeString(font.family);

		if (!fontMap.has(key)) {
			fontMap.set(key, []);
		}

		fontMap.get(key).push(font);
	}

	fontCache = {
		system: new Map(),
		adobe: new Map([
			["all", new Map()],
			["other", new Map()]
		]),
		custom: new Map(),
		all: new Map(fontMap)
	};

	var count = 0;

	// Анализ шрифтов Adobe
	for (const folder of adobeFolders) {
		if (folder != "" && !fontCache.adobe.has(folder)) {
			fontCache.adobe.set(folder, new Map());
		}

		for (const file of await getFontFiles(FONT_PATHS.adobe + "/" + folder)) {
			const fontName = deepNormalizeString(await getFontNameFromFile(FONT_PATHS.adobe + "/" + folder + "/" + file));
			if (fontName) {
				var psFont = fontMap.get(fontName);
				if (psFont) {
					fontCache.adobe.get("all").set(fontName, psFont);
					if (fontCache.adobe.has(folder)) {
						fontCache.adobe.get(folder).set(fontName, psFont);
					} else {
						fontCache.adobe.get("other").set(fontName, psFont);
					}
					count += psFont.length;
					fontMap.delete(fontName);
				}
			}
		}
	}

	// Анализ системных шрифтов
	for (const file of await getFontFiles(FONT_PATHS.system)) {
		const fontName = deepNormalizeString(await getFontNameFromFile(FONT_PATHS.system + "/" + file));
		if (fontName) {
			var psFont = fontMap.get(fontName);
			if (psFont) {
				fontCache.system.set(fontName, psFont);
				fontMap.delete(fontName);
			}
		}
	}

	// Остальные шрифты
	fontCache.custom = fontMap;
}

async function showAlert(message) {
	await app.showAlert(message);
}

// DOM-элементы
const refreshButton = document.getElementById("refreshButton");
const infoButton = document.getElementById("infoButton");
const searchInput = document.getElementById("search");
const tabs = document.querySelectorAll(".tab");
const fontList = document.getElementById("fontList");
const closeButton = document.getElementById("closeButton");
const typeList = document.getElementById("typeList");
const folderButton = document.getElementById("folderButton");

// Применение шрифта
async function applyFont(font) {
	try {
		if (!app.activeDocument) {
			throw new Error("Откройте документ в Photoshop!");
		}

		const activeLayer = app.activeDocument.activeLayers[0];
		if (!activeLayer) {
			throw new Error("Выберите слой! (" + activeLayer + ")");
		}

		if (activeLayer.kind !== "text") {
			throw new Error("Выберите текстовый слой! (" + activeLayer.kind + ")");
		}

		await core.executeAsModal(async () => {
			activeLayer.textItem.characterStyle.font = font.postScriptName || font.name;
		}, { commandName: "Apply Font" });
	} catch (error) {
		showAlert(error.message);
		console.error("Ошибка:", error);
	}
}

var cat = "all";
var customType = "all";

// Функция для проверки доступности шрифта
function isFontAvailable(fontName, callback) {
	const span = document.createElement('span');
	span.style.fontSize = '72px';
	span.style.position = 'absolute';
	span.style.left = '-9999px';
	span.style.fontFamily = 'sans-serif';
	span.textContent = 'abcdefghijklmnopqrstuvwxyz';

	document.body.appendChild(span);
	const baseWidth = span.offsetWidth;

	span.style.fontFamily = `"${fontName}", sans-serif`;

	// Даем время на применение стилей
	setTimeout(() => {
		const testWidth = span.offsetWidth;
		document.body.removeChild(span);
		callback(testWidth !== baseWidth);
	}, 100);
}

// Отрисовка списка шрифтов
async function renderFonts(searchTerm = "", category = cat) {
	fontList.innerHTML = "";
	typeList.innerHTML = "";
	let fontsToShow = new Map();
	if (category === "all") fontsToShow = fontCache.all;
	if (category === "system") fontsToShow = fontCache.system;
	if (category === "adobe") {
		fontCache.adobe.forEach((fonts, type) => {
			const tab = document.createElement("h3");
			tab.className = "type-tab";
			tab.dataset.category = type;
			tab.textContent = `${tabNames.get(type) ?? type} (${fonts.size})`;
			if (type == customType) {
				tab.classList.add("active");
			}
			tab.onclick = () => {
				typeList.children.forEach(t => t.classList.remove("active"));
				tab.classList.add("active");
				customType = tab.dataset.category;
				renderFonts(searchInput.value);
			}
			typeList.appendChild(tab);
		});
		fontsToShow = fontCache.adobe.get(customType);
	}
	if (category === "custom") fontsToShow = fontCache.custom;

	if (searchTerm) {
		searchTerm = searchTerm.toLowerCase()
		fontsToShow = [...fontsToShow].reduce((result, [family, fonts]) => {
			fonts.forEach(font => {
				if (family.toLowerCase().includes(searchTerm) || font.postScriptName.toLowerCase().includes(searchTerm) || font.name.toLowerCase().includes(searchTerm)) {
					if (!result.has(family)) {
						result.set(family, []);
					}
					result.get(family).push(font);
				}
			});
			return result;
		}, new Map())
	}

	if (fontsToShow.length === 0) {
		const emptyMsg = document.createElement("div");
		emptyMsg.textContent = "Шрифты не найдены";
		emptyMsg.style.padding = "10px";
		emptyMsg.style.color = "#999";
		fontList.appendChild(emptyMsg);
		return;
	}

	// Отрисовка
	fontsToShow.forEach((fonts, family) => {
		const familyItem = document.createElement("div");
		familyItem.className = "family-item";
		familyItem.onclick = () => applyFont(fonts[0]);

		const familyNameBar = document.createElement("div");
		familyNameBar.className = "family-header"

		const familyLeft = document.createElement("div");
		familyLeft.className = "family-left";

		const openButton = document.createElement("button");
		openButton.className = "family-open";
		openButton.textContent = "▼"

		const familyNameItem = document.createElement("h6");
		familyNameItem.textContent = fonts[0].family;

		const previewItem = document.createElement("div");
		previewItem.textContent = "Sample";

		previewItem.style.fontFamily = `"${deepNormalizeString(fonts[0].family)}"`;

		// Максимальное количество времени на проверку шрифта
		const maxNameTimeout = 1000;
		// Время обновления проверки (каждые updateTime мс)
		const updateTime = 100;
		// Количество обновлений
		const updateCount = maxNameTimeout / updateTime;

		var lastCallbackCount = 0;
		var firstCallbackCount = 0;

		const lastCallback = () => {
			const w = previewItem.offsetWidth;
			const h = previewItem.offsetHeight;

			if (w == 79 && h == 32) {
				familyItem.className = "family-item-bad";
				if (lastCallbackCount < updateCount) {
					lastCallbackCount++;
					setTimeout(lastCallback, updateTime);
				} else {
					previewItem.style.fontFamily = `"${deepNormalizeString(fonts[0].postScriptName)}"`;
				}
			} else {
				familyItem.className = "family-item";
			}
		};

		const firstCallback = () => {
			const w = previewItem.offsetWidth;
			const h = previewItem.offsetHeight;

			if (w == 0 || h == 0) {
				if (firstCallbackCount < updateCount) {
					firstCallbackCount++;
					setTimeout(firstCallback, updateTime);
				}
			} else if (w == 79 && h == 32) {
				previewItem.style.fontFamily = `"${deepNormalizeString(fonts[0].name)}"`;
				setTimeout(lastCallback, updateTime);
			}
		};

		setTimeout(firstCallback, updateTime);

		familyItem.addEventListener('contextmenu', function (ev) {
			ev.preventDefault();

			switch (previewItem.dataset.category) {
				case "name":
					previewItem.style.fontFamily = `"${deepNormalizeString(fonts[0].postScriptName)}"`;
					previewItem.dataset.category = "postScriptName";
					break;
				case "postScriptName":
					previewItem.style.fontFamily = `"${deepNormalizeString(fonts[0].family)}"`;
					previewItem.dataset.category = "family";
					break;
				case "family":
				default:
					previewItem.style.fontFamily = `"${deepNormalizeString(fonts[0].name)}"`;
					previewItem.dataset.category = "name";
					break;
			}



			return false;
		}, false);

		previewItem.style.fontSize = '24px';

		familyLeft.appendChild(openButton);
		familyLeft.appendChild(familyNameItem);
		familyNameBar.appendChild(familyLeft);
		familyNameBar.appendChild(previewItem);
		familyItem.appendChild(familyNameBar);

		if (fonts.length > 1) {
			const fontsContainerItem = document.createElement("div");
			fontsContainerItem.style.display = 'none';

			openButton.onclick = (event) => {
				event.stopPropagation();
				if (fontsContainerItem.style.display === 'none') {
					fontsContainerItem.style.display = 'block';
					openButton.textContent = "▲"
				} else {
					fontsContainerItem.style.display = 'none';
					openButton.textContent = "▼"
				}
			}

			fonts.forEach(font => {
				const fontItem = document.createElement("h7");
				fontItem.className = "font-item";
				fontItem.textContent = font.name + " : " + font.postScriptName;
				fontItem.onclick = (event) => {
					event.stopPropagation();
					applyFont(font);
				}
				fontsContainerItem.appendChild(fontItem);
			});

			familyItem.appendChild(fontsContainerItem);
		} else {
			if (familyNameItem.textContent != fonts[0].name) {
				familyNameItem.textContent += " (" + fonts[0].name + ")";
			}
			openButton.style.visibility = "hidden";
		}
		fontList.appendChild(familyItem);
	});
}

function setFolderPath(path) {
	FONT_PATHS.adobe = path;
}

// Функция для нормализации пути (замена \ на /)
function normalizePath(path) {
	if (!path) return path;
	// Заменяем все обратные слеши на прямые
	return path.replace(/\\/g, '/');
}

// Функция для загрузки пути из настроек (localStorage)
function loadSavedFolderPath() {
	const savedPath = localStorage.getItem("fonts-folder");

	if (savedPath) {
		console.warn("Загружен сохраненный путь:", savedPath);
		return savedPath;
	} else {
		console.warn("Не загружен сохраненный путь");
		return "file:C:/Program Files/Common Files/Adobe/Fonts";
	}
}

// Функция для сохранения пути в настройки (localStorage)
function saveFolderPath(path) {
	if (path) {
		localStorage.setItem("fonts-folder", path);
		console.warn("Путь сохранен:", path);
		// Обновляем отображение
		setFolderPath(path);
	}
}

// Асинхронная функция для выбора папки
async function handleSelectFolder() {
	try {
		// Запрашиваем у пользователя папку через системный диалог.
		// Функция getFolder() возвращает объект Folder, если пользователь выбрал папку,
		// или null, если диалог был отменен.
		const selectedFolder = await uxp.storage.localFileSystem.getFolder();

		if (selectedFolder) {
			// Получаем "родной" путь к папке (он может быть в URN-формате)
			// Для отображения пользователю лучше использовать .nativePath
			const folderPath = "file:" + normalizePath(selectedFolder.nativePath);

			// Сохраняем путь в настройки
			saveFolderPath(folderPath);

			// Можно также показать временное сообщение (опционально)
			showAlert(`Выбрана папка: ${folderPath}`);
			init();
		} else {
			console.warn("Пользователь отменил выбор папки.");
		}
	} catch (error) {
		console.error("Ошибка при выборе папки:", error);
		showAlert("Ошибка доступа к файловой системе.");
	}
}

// Инициализация
async function init() {
	try {
		setFolderPath(loadSavedFolderPath());

		await analyzeFonts();

		// Обновление списка
		refreshButton.onclick = () => init();

		infoButton.onclick = () => {
			showAlert(`Шрифты Adobe должны быть размещены по этому пути:\n${FONT_PATHS.adobe}\n\nдля правильного отображения установить шрифты\nили разместить по этому пути: ${FONT_PATHS.system}\n\nпосле удаления шрифта из Adobe\nне забыть удалить шрифт из системы`)
		};

		folderButton.onclick = () => {
			handleSelectFolder()
		};

		searchInput.oninput = (e) => renderFonts(e.target.value);

		closeButton.onclick = () => {
			searchInput.value = "";
			renderFonts();
		}

		tabs.forEach(tab => {
			if (tab.dataset.category == "all") tab.innerHTML = tabNames.get(tab.dataset.category) + " (" + fontCache.all.size + ")";
			if (tab.dataset.category == "system") tab.innerHTML = tabNames.get(tab.dataset.category) + " (" + fontCache.system.size + ")";
			if (tab.dataset.category == "adobe") tab.innerHTML = tabNames.get(tab.dataset.category) + " (" + fontCache.adobe.get("all").size + ")";
			if (tab.dataset.category == "custom") tab.innerHTML = tabNames.get(tab.dataset.category) + " (" + fontCache.custom.size + ")";
			tab.addEventListener("click", () => {
				tabs.forEach(t => t.classList.remove("active"));
				tab.classList.add("active");
				cat = tab.dataset.category;
				if (tab.dataset.category == "adobe") {
					typeList.style.display = 'block';
				} else {
					typeList.style.display = 'none';
				}
				renderFonts(searchInput.value, tab.dataset.category);
			});

			if (tab.classList.contains('active')) {
				cat = tab.dataset.category;
				if (tab.dataset.category == "adobe") {
					typeList.style.display = 'block';
				} else {
					typeList.style.display = 'none';
				}
			}
		});

		renderFonts(searchInput.value);

	} catch (error) {
		showAlert("Ошибка инициализации: " + error.message);
	}
}

// Запуск
init();
import * as fs from 'fs';
import * as path from 'path';
import * as ts from 'typescript';

// 1. Mock Google Apps Script Globals
(global as any).Logger = { ...console, log: jest.fn() };
(global as any).console = console;

(global as any).PropertiesService = {
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    getProperties: jest.fn(() => ({})),
  })),
};

(global as any).UrlFetchApp = {
  fetch: jest.fn(),
};

(global as any).SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(),
  newConditionalFormatRule: jest.fn(() => ({
    whenFormulaSatisfied: jest.fn().mockReturnThis(),
    whenTextEqualTo: jest.fn().mockReturnThis(),
    setBackground: jest.fn().mockReturnThis(),
    setFontColor: jest.fn().mockReturnThis(),
    setRanges: jest.fn().mockReturnThis(),
    build: jest.fn(),
  })),
  newDataValidation: jest.fn(() => ({
    requireValueInList: jest.fn().mockReturnThis(),
    build: jest.fn(),
  })),
  BorderStyle: { SOLID_MEDIUM: 'SOLID_MEDIUM' },
  ProtectionType: { RANGE: 'RANGE' },
};

(global as any).Utilities = {
  formatDate: jest.fn(),
  sleep: jest.fn(),
  base64Encode: jest.fn((str: string) => Buffer.from(str).toString('base64')),
};

(global as any).Session = {
  getActiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'admin@example.com')
  })),
};

// 2. Load GAS Source into Global Scope
const srcDir = path.resolve(__dirname, '../src');
const sourceFiles = ['Config.ts', 'Http.ts', 'Slack.ts', 'Spreadsheet.ts', 'Discord.ts', 'Menu.ts'];

let allCode = '';
for (const file of sourceFiles) {
  allCode += fs.readFileSync(path.join(srcDir, file), 'utf8') + '\n';
}

// Extract top-level symbol names
const exportedNames = new Set<string>();
const functionRegex = /^function\s+([a-zA-Z0-9_]+)\s*\(/gm;
const constRegex = /^(?:const|let|var)\s+([a-zA-Z0-9_]+)\s*(?:=|:)/gm;

let match;
while ((match = functionRegex.exec(allCode)) !== null) exportedNames.add(match[1]);
while ((match = constRegex.exec(allCode)) !== null) exportedNames.add(match[1]);

// Transpile the code
const transpiled = ts.transpileModule(allCode, {
  compilerOptions: { target: ts.ScriptTarget.ES2019, module: ts.ModuleKind.None }
}).outputText;

const exportStatements = Array.from(exportedNames)
  .map(name => `
    try {
      Object.defineProperty(_exports, "${name}", {
        get: () => typeof ${name} !== 'undefined' ? ${name} : undefined,
        set: (val) => { try { ${name} = val; } catch(e) {} },
        configurable: true,
        enumerable: true
      });
    } catch(e) {}
  `)
  .join('\n');

const wrapperCode = `
  ${transpiled}
  ${exportStatements}
`;

const _exports = {};
const wrapper = new Function('global', '_exports', wrapperCode);
wrapper(global, _exports);

// Attach everything to global preserving getters/setters
for (const key of Object.keys(_exports)) {
  const descriptor = Object.getOwnPropertyDescriptor(_exports, key);
  if (descriptor) {
    Object.defineProperty(global, key, descriptor);
  }
}

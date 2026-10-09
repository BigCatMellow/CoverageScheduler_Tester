'use strict';

const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const gasDir = path.join(root, 'google-apps-script');
const deployDir = path.join(root, 'apps-script-deploy');

const SERVER_ORDER = [
  'coverage-foundation.gs',
  'setup.gs',
  'teacher-schedule-adapter.gs',
  'class-schedule.gs',
  'field-trips.gs',
  'schedule-cache.gs',
  'coverage-staff.gs',
  'field-trip-pool.gs',
  'manual-coverage.gs',
  'coverage-assignment.gs',
  'scheduler.gs',
  'coverage-save.gs',
  'web-ui-data.gs',
  'field-trip-ui.gs',
  'handout.gs',
  'code.gs'
];

const UI_PARTIALS = [
  'styles',
  'ui-core',
  'ui-plan',
  'ui-people',
  'ui-manual',
  'ui-actions'
];

function readGas(name) {
  return fs.readFileSync(path.join(gasDir, name), 'utf8').replace(/\r\n/g, '\n');
}

function assertServerSet() {
  const discovered = fs.readdirSync(gasDir)
    .filter(name => name.endsWith('.gs'))
    .sort();
  const expected = SERVER_ORDER.slice().sort();
  if (JSON.stringify(discovered) !== JSON.stringify(expected)) {
    throw new Error(
      'Apps Script server file list changed. Update SERVER_ORDER.\n' +
      'Discovered: ' + discovered.join(', ') + '\n' +
      'Expected: ' + expected.join(', ')
    );
  }
}

function buildServerBundle() {
  assertServerSet();
  return SERVER_ORDER.map(name =>
    '// ============================================================================\n' +
    '// SOURCE: google-apps-script/' + name + '\n' +
    '// ============================================================================\n\n' +
    readGas(name).trimEnd()
  ).join('\n\n') + '\n';
}

function buildIndexBundle() {
  let index = readGas('index.html');
  for (const name of UI_PARTIALS) {
    const marker = "<?!= includeCoveragePartial_('" + name + "'); ?>";
    if (!index.includes(marker)) {
      throw new Error('Missing index partial marker: ' + marker);
    }
    index = index.replace(marker, readGas(name + '.html').trimEnd());
  }
  if (index.includes('<?!= includeCoveragePartial_')) {
    throw new Error('Unresolved UI partial remains in bundled index.html');
  }
  return index;
}

function buildManifest() {
  return readGas('appsscript.json');
}

function buildReadme() {
  return [
    'Coverage Scheduler — Manual Apps Script Deployment',
    '',
    'You only need to replace three files in Apps Script:',
    '',
    '1. CoverageScheduler.gs',
    '2. index.html (the Apps Script HTML file should be named index)',
    '3. appsscript.json',
    '',
    'IMPORTANT:',
    '- Once you switch to this bundle, remove the old modular .gs files and UI partial HTML files from the Apps Script project.',
    '- Leaving both the modular files and CoverageScheduler.gs in the same Apps Script project will define functions/constants twice.',
    '- Keep index.html as the only full-page web-app HTML file.',
    '',
    'Then save and use:',
    'Deploy -> Manage deployments -> Edit -> New version -> Deploy',
    ''
  ].join('\n');
}

function buildAll() {
  return {
    'CoverageScheduler.gs': buildServerBundle(),
    'index.html': buildIndexBundle(),
    'appsscript.json': buildManifest(),
    'README.txt': buildReadme()
  };
}

function writeAll() {
  fs.rmSync(deployDir, { recursive: true, force: true });
  fs.mkdirSync(deployDir, { recursive: true });
  const outputs = buildAll();
  for (const [name, content] of Object.entries(outputs)) {
    fs.writeFileSync(path.join(deployDir, name), content, 'utf8');
  }
  return outputs;
}

if (require.main === module) {
  const outputs = writeAll();
  console.log(
    'Built manual Apps Script bundle: ' +
    Object.keys(outputs).join(', ')
  );
}

module.exports = {
  SERVER_ORDER,
  UI_PARTIALS,
  buildServerBundle,
  buildIndexBundle,
  buildManifest,
  buildReadme,
  buildAll,
  writeAll
};

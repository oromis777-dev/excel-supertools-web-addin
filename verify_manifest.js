const fs = require('fs');
const xml = fs.readFileSync('manifest.xml', 'utf8');

console.log('--- Checking manifest.xml ---');
console.log('Length:', xml.length);

// Extract all resid
const resourceRefs = [];
const refRegex = /resid="([^"]+)"/gi;
let m;
while ((m = refRegex.exec(xml)) !== null) {
  resourceRefs.push(m[1]);
}
console.log('Resource references found in controls:', resourceRefs);

// Extract defined resource IDs in bt:Images, bt:Urls, bt:ShortStrings, bt:LongStrings
const definedResources = [];
const defRegex = /<(?:bt:(?:Image|Url|String))\s+id="([^"]+)"/gi;
while ((m = defRegex.exec(xml)) !== null) {
  definedResources.push(m[1]);
}
console.log('Defined resources in <Resources>:', definedResources);

// Check if any reference is missing
const missing = resourceRefs.filter(r => !definedResources.includes(r));
if (missing.length > 0) {
  console.error('ERROR: Missing resource definitions:', missing);
} else {
  console.log('SUCCESS: All resource references exist in <Resources>!');
}

// Check Function Names in ExecuteFunction
const funcRefs = [];
const funcRegex = /<FunctionName>([^<]+)<\/FunctionName>/g;
while ((m = funcRegex.exec(xml)) !== null) {
  funcRefs.push(m[1]);
}
console.log('Function names referenced:', funcRefs);

// Check if commands.js defines them
const commandsJs = fs.readFileSync('src/commands/commands.js', 'utf8');
const missingFuncs = funcRefs.filter(f => !commandsJs.includes(f));
if (missingFuncs.length > 0) {
  console.error('ERROR: Missing functions in commands.js:', missingFuncs);
} else {
  console.log('SUCCESS: All manifest functions exist in commands.js!');
}

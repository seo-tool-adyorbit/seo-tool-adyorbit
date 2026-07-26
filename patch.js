const fs = require('fs');
const file = '/Users/rvisharma/TFN changer Tool/js/app.js';
let content = fs.readFileSync(file, 'utf8');

const oldBuildReplacementString = `    // Helper: The Magic Replacement String builder
    const buildReplacementString = (matchStr, targetChars) => {
        // 1. Count original digits
        let numSourceDigits = 0;
        for (let char of matchStr) {
            const normalized = char.normalize('NFKC');
            if (/^\\d$/.test(normalized)) {
                numSourceDigits++;
            }
        }

        const numTargetDigits = targetChars.length;
        // If source has more digits (e.g. 1-800 vs 800), skip replacing the leading digits (keep country code)
        let digitsToSkip = Math.max(0, numSourceDigits - numTargetDigits);

        let targetIdx = 0;
        let result = '';
        let lastOriginalDigitChar = null;
        
        for (let char of matchStr) {
            const normalized = char.normalize('NFKC');
            if (/^\\d$/.test(normalized)) {
                lastOriginalDigitChar = char;
                if (digitsToSkip > 0) {
                    result += char; // Keep original digit unchanged (e.g. leading 1)
                    digitsToSkip--;
                } else if (targetIdx < numTargetDigits) {
                    result += formatDigitLike(char, targetChars[targetIdx]);
                    targetIdx++;
                }
            } else {
                result += char; // Keep symbols, emojis, formatting chars
            }
        }
        
        while (targetIdx < numTargetDigits) {
            const targetChar = targetChars[targetIdx];
            if (lastOriginalDigitChar) {
                result += formatDigitLike(lastOriginalDigitChar, targetChar);
            } else {
                result += targetChar;
            }
            targetIdx++;
        }
        return result;
    };`;

const newBuildReplacementString = `    // Helper: The Magic Replacement String builder
    const buildReplacementString = (matchStr, targetNumberStr) => {
        let startIndex = -1;
        let sampleOriginalDigitChar = null;
        
        // Find leading context (e.g. spaces, colons before the actual number starts)
        for (let i = 0; i < matchStr.length; i++) {
            const char = matchStr[i];
            const normalized = char.normalize('NFKC');
            
            if (!sampleOriginalDigitChar && /^\\d$/.test(normalized)) {
                sampleOriginalDigitChar = char;
            }
            
            if (startIndex === -1 && (/^\\d$/.test(normalized) || char === '+' || char === '(' || char === '[')) {
                startIndex = i;
            }
        }
        
        if (startIndex === -1) startIndex = 0;
        
        const leadingContext = matchStr.substring(0, startIndex);
        let result = leadingContext;
        
        // Append the target number using its own formatting, 
        // but preserve the unicode font styling of the original digits
        for (let char of targetNumberStr) {
            if (/^\\d$/.test(char.normalize('NFKC'))) {
                if (sampleOriginalDigitChar) {
                    result += formatDigitLike(sampleOriginalDigitChar, char);
                } else {
                    result += char;
                }
            } else {
                result += char; // Keep target's formatting symbols (+, -, space, etc)
            }
        }
        
        return result;
    };`;

content = content.replace(oldBuildReplacementString, newBuildReplacementString);

// Also need to fix the caller to pass targetNumberStr instead of targetChars
const oldCaller = `const replacementText = buildReplacementString(m.text, targetChars);`;
const newCaller = `const replacementText = buildReplacementString(m.text, targetNumberStr);`;
content = content.replace(oldCaller, newCaller);

fs.writeFileSync(file, content);
console.log("Patched successfully");

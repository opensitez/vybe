/// String algorithms — classic CS string problems
use super::helpers::run_js;

#[test]
fn reverse_words() {
    assert_eq!(
        run_js(
            r#"
function reverseWords(s) {
    return s.trim().split(/\s+/).reverse().join(" ");
}
console.log(reverseWords("Hello World"));
console.log(reverseWords("  the sky is blue  "));
"#
        ),
        vec!["World Hello", "blue is sky the"]
    );
}

#[test]
fn is_palindrome_string() {
    assert_eq!(
        run_js(
            r#"
function isPalindrome(s) {
    const cleaned = s.toLowerCase().replace(/[^a-z0-9]/g, "");
    return cleaned === cleaned.split("").reverse().join("");
}
console.log(isPalindrome("A man, a plan, a canal: Panama"));
console.log(isPalindrome("race a car"));
console.log(isPalindrome("Was it a car or a cat I saw?"));
"#
        ),
        vec!["true", "false", "true"]
    );
}

#[test]
fn count_vowels() {
    assert_eq!(
        run_js(
            r#"
const countVowels = s => (s.match(/[aeiouAEIOU]/g) || []).length;
console.log(countVowels("Hello World"));
console.log(countVowels("rhythm"));
console.log(countVowels("aeiou"));
"#
        ),
        vec!["3", "0", "5"]
    );
}

#[test]
fn title_case() {
    assert_eq!(
        run_js(
            r#"
const titleCase = s => s.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
console.log(titleCase("hello world"));
console.log(titleCase("the quick brown fox"));
"#
        ),
        vec!["Hello World", "The Quick Brown Fox"]
    );
}

#[test]
fn compress_string() {
    assert_eq!(
        run_js(
            r#"
function compress(s) {
    let result = "";
    let i = 0;
    while (i < s.length) {
        let j = i;
        while (j < s.length && s[j] === s[i]) j++;
        result += s[i] + (j - i > 1 ? (j - i) : "");
        i = j;
    }
    return result.length < s.length ? result : s;
}
console.log(compress("aabcccdddd"));
console.log(compress("abc"));
"#
        ),
        vec!["a2bc3d4", "abc"]
    );
}

#[test]
fn find_all_substrings() {
    assert_eq!(
        run_js(
            r#"
function findAll(text, pattern) {
    const indices = [];
    let idx = text.indexOf(pattern);
    while (idx !== -1) {
        indices.push(idx);
        idx = text.indexOf(pattern, idx + 1);
    }
    return indices;
}
console.log(findAll("abababab", "ab").join(","));
console.log(findAll("hello", "xyz").join(","));
"#
        ),
        vec!["0,2,4,6", ""]
    );
}

#[test]
fn string_interleave() {
    assert_eq!(
        run_js(
            r#"
function interleave(a, b) {
    const result = [];
    const len = Math.max(a.length, b.length);
    for (let i = 0; i < len; i++) {
        if (i < a.length) result.push(a[i]);
        if (i < b.length) result.push(b[i]);
    }
    return result.join("");
}
console.log(interleave("abc", "12345"));
console.log(interleave("xyz", ""));
"#
        ),
        vec!["a1b2c345", "xyz"]
    );
}

#[test]
fn roman_to_int() {
    assert_eq!(
        run_js(
            r#"
function romanToInt(s) {
    const map = { I:1, V:5, X:10, L:50, C:100, D:500, M:1000 };
    let result = 0;
    for (let i = 0; i < s.length; i++) {
        const curr = map[s[i]], next = map[s[i+1]];
        result += (next > curr) ? -curr : curr;
    }
    return result;
}
console.log(romanToInt("III"));
console.log(romanToInt("IV"));
console.log(romanToInt("MCMXCIV"));
"#
        ),
        vec!["3", "4", "1994"]
    );
}

#[test]
fn int_to_roman() {
    assert_eq!(
        run_js(
            r#"
function intToRoman(num) {
    const vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
    const syms = ["M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I"];
    let result = "";
    for (let i = 0; i < vals.length; i++) {
        while (num >= vals[i]) { result += syms[i]; num -= vals[i]; }
    }
    return result;
}
console.log(intToRoman(3));
console.log(intToRoman(58));
console.log(intToRoman(1994));
"#
        ),
        vec!["III", "LVIII", "MCMXCIV"]
    );
}

#[test]
fn tokenize_expression() {
    assert_eq!(
        run_js(
            r#"
function tokenize(expr) {
    return expr.match(/\d+|[+\-*/()]/g) || [];
}
console.log(tokenize("3+4*(2-1)").join(","));
console.log(tokenize("100/25+5").join(","));
"#
        ),
        vec!["3,+,4,*,(,2,-,1,)", "100,/,25,+,5"]
    );
}

#[test]
fn string_rotate() {
    assert_eq!(
        run_js(
            r#"
const rotateStr = (s, n) => s.slice(n % s.length) + s.slice(0, n % s.length);
console.log(rotateStr("abcde", 2));
console.log(rotateStr("hello", 0));
console.log(rotateStr("abcde", 5));
"#
        ),
        vec!["cdeab", "hello", "abcde"]
    );
}

#[test]
fn most_common_char() {
    assert_eq!(
        run_js(
            r#"
function mostCommon(s) {
    const freq = {};
    for (const c of s) freq[c] = (freq[c] ?? 0) + 1;
    return Object.entries(freq).sort((a, b) => b[1] - a[1])[0][0];
}
console.log(mostCommon("aabbccddeeee"));
console.log(mostCommon("hello"));
"#
        ),
        vec!["e", "l"]
    );
}

#[test]
fn zigzag_string_pattern() {
    assert_eq!(
        run_js(
            r#"
function zigzag(s, rows) {
    if (rows === 1) return s;
    const buckets = Array.from({length: rows}, () => "");
    let row = 0, dir = 1;
    for (const c of s) {
        buckets[row] += c;
        if (row === 0) dir = 1;
        else if (row === rows - 1) dir = -1;
        row += dir;
    }
    return buckets.join("");
}
console.log(zigzag("PAYPALISHIRING", 3));
console.log(zigzag("AB", 1));
"#
        ),
        vec!["PAHNAPLSIIGYIR", "AB"]
    );
}

#[test]
fn bracket_matching() {
    assert_eq!(
        run_js(
            r#"
function isValid(s) {
    const stack = [];
    const map = { ')':'(', ']':'[', '}':'{' };
    for (const c of s) {
        if ("([{".includes(c)) stack.push(c);
        else if (stack.pop() !== map[c]) return false;
    }
    return stack.length === 0;
}
console.log(isValid("()[]{}"));
console.log(isValid("([)]"));
console.log(isValid("{[]}"));
"#
        ),
        vec!["true", "false", "true"]
    );
}

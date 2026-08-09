/// Math and number utility patterns
use super::helpers::run_js;

#[test]
fn gcd_lcm() {
    assert_eq!(
        run_js(
            r#"
function gcd(a, b) { return b === 0 ? a : gcd(b, a % b); }
function lcm(a, b) { return (a / gcd(a, b)) * b; }
console.log(gcd(48, 18));
console.log(lcm(4, 6));
console.log(gcd(17, 5));
"#
        ),
        vec!["6", "12", "1"]
    );
}

#[test]
fn sieve_of_eratosthenes() {
    assert_eq!(
        run_js(
            r#"
function sieve(n) {
    const primes = Array(n + 1).fill(true);
    primes[0] = primes[1] = false;
    for (let i = 2; i * i <= n; i++) {
        if (primes[i]) for (let j = i*i; j <= n; j+=i) primes[j] = false;
    }
    return primes.map((p,i)=>p?i:-1).filter(n=>n>0);
}
const p = sieve(30);
console.log(p.join(","));
"#
        ),
        vec!["2,3,5,7,11,13,17,19,23,29"]
    );
}

#[test]
fn prime_factorization() {
    assert_eq!(
        run_js(
            r#"
function primeFactors(n) {
    const factors = [];
    for (let d = 2; d * d <= n; d++) {
        while (n % d === 0) { factors.push(d); n /= d; }
    }
    if (n > 1) factors.push(n);
    return factors;
}
console.log(primeFactors(12).join(","));
console.log(primeFactors(100).join(","));
console.log(primeFactors(17).join(","));
"#
        ),
        vec!["2,2,3", "2,2,5,5", "17"]
    );
}

#[test]
fn power_with_modulo() {
    assert_eq!(
        run_js(
            r#"
function powMod(base, exp, mod) {
    let result = 1n;
    base = BigInt(base) % BigInt(mod);
    exp = BigInt(exp);
    mod = BigInt(mod);
    while (exp > 0n) {
        if (exp % 2n === 1n) result = result * base % mod;
        exp >>= 1n;
        base = base * base % mod;
    }
    return Number(result);
}
console.log(powMod(2, 10, 1000));
console.log(powMod(3, 5, 13));
"#
        ),
        vec!["24", "9"]
    );
}

#[test]
fn sum_of_digits() {
    assert_eq!(
        run_js(
            r#"
function digitSum(n) {
    return Math.abs(n).toString().split("").reduce((s, d) => s + Number(d), 0);
}
function digitalRoot(n) {
    while (n > 9) n = digitSum(n);
    return n;
}
console.log(digitSum(12345));
console.log(digitalRoot(9875));
console.log(digitalRoot(0));
"#
        ),
        vec!["15", "2", "0"]
    );
}

#[test]
fn factorial_and_permutations() {
    assert_eq!(
        run_js(
            r#"
function factorial(n) { return n <= 1 ? 1 : n * factorial(n - 1); }
function permutations(n, r) { return factorial(n) / factorial(n - r); }
function combinations(n, r) { return permutations(n, r) / factorial(r); }
console.log(factorial(5));
console.log(permutations(5, 2));
console.log(combinations(5, 2));
"#
        ),
        vec!["120", "20", "10"]
    );
}

#[test]
fn number_base_conversions() {
    assert_eq!(
        run_js(
            r#"
const dec = n => parseInt(n, 2);  // binary to decimal
const hex = n => n.toString(16);
const bin = n => n.toString(2);
console.log(dec("1010"));
console.log(hex(255));
console.log(bin(42));
console.log(parseInt("ff", 16));
"#
        ),
        vec!["10", "ff", "101010", "255"]
    );
}

#[test]
fn is_perfect_number() {
    assert_eq!(
        run_js(
            r#"
function isPerfect(n) {
    if (n <= 1) return false;
    let sum = 1;
    for (let i = 2; i * i <= n; i++) {
        if (n % i === 0) { sum += i; if (i !== n/i) sum += n/i; }
    }
    return sum === n;
}
console.log(isPerfect(6));
console.log(isPerfect(28));
console.log(isPerfect(12));
"#
        ),
        vec!["true", "true", "false"]
    );
}

#[test]
fn statistics_mean_median_mode() {
    assert_eq!(
        run_js(
            r#"
function mean(arr) { return arr.reduce((a,b)=>a+b,0) / arr.length; }
function median(arr) {
    const s = [...arr].sort((a,b)=>a-b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m-1]+s[m]) / 2;
}
function mode(arr) {
    const freq = new Map();
    for (const x of arr) freq.set(x, (freq.get(x)??0)+1);
    return [...freq.entries()].sort((a,b)=>b[1]-a[1])[0][0];
}
const data = [1, 2, 2, 3, 4, 4, 4, 5];
console.log(mean(data));
console.log(median(data));
console.log(mode(data));
"#
        ),
        vec!["3.125", "3.5", "4"]
    );
}

#[test]
fn clamp_lerp_normalize() {
    assert_eq!(
        run_js(
            r#"
const clamp = (v, min, max) => Math.max(min, Math.min(max, v));
const lerp = (a, b, t) => a + (b - a) * t;
const normalize = (v, min, max) => (v - min) / (max - min);
console.log(clamp(150, 0, 100));
console.log(lerp(0, 100, 0.5));
console.log(normalize(75, 0, 100));
console.log(clamp(-5, 0, 100));
"#
        ),
        vec!["100", "50", "0.75", "0"]
    );
}

#[test]
fn number_palindrome() {
    assert_eq!(
        run_js(
            r#"
function isPalindromeNum(n) {
    if (n < 0) return false;
    const s = n.toString();
    return s === s.split("").reverse().join("");
}
console.log(isPalindromeNum(121));
console.log(isPalindromeNum(-121));
console.log(isPalindromeNum(1001));
console.log(isPalindromeNum(10));
"#
        ),
        vec!["true", "false", "true", "false"]
    );
}

#[test]
fn fibonacci_sequence_generator() {
    assert_eq!(
        run_js(
            r#"
function* fibSeq() {
    let [a, b] = [0, 1];
    while (true) { yield a; [a, b] = [b, a+b]; }
}
const gen = fibSeq();
const first10 = Array.from({length: 10}, () => gen.next().value);
console.log(first10.join(","));
"#
        ),
        vec!["0,1,1,2,3,5,8,13,21,34"]
    );
}

#[test]
fn integer_square_root() {
    assert_eq!(
        run_js(
            r#"
function isqrt(n) {
    if (n < 0) return NaN;
    let x = Math.floor(Math.sqrt(n));
    while (x * x > n) x--;
    while ((x+1)*(x+1) <= n) x++;
    return x;
}
console.log(isqrt(16));
console.log(isqrt(17));
console.log(isqrt(100));
console.log(isqrt(0));
"#
        ),
        vec!["4", "4", "10", "0"]
    );
}

#[test]
fn vector_normalization_hypot() {
    assert_eq!(
        run_js(
            r#"
const norm = (x, y) => { const len = Math.hypot(x, y); return [x / len, y / len]; };
console.log(norm(3, 4).join(","));
"#
        ),
        vec!["0.6,0.8"]
    );
}

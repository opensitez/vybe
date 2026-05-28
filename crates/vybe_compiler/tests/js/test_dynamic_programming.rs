/// Dynamic programming patterns in JS

use super::helpers::run_js;

#[test]
fn fibonacci_dp() {
    assert_eq!(run_js(r#"
function fib(n) {
    if (n <= 1) return n;
    let a = 0, b = 1;
    for (let i = 2; i <= n; i++) [a, b] = [b, a + b];
    return b;
}
console.log(fib(0));
console.log(fib(1));
console.log(fib(10));
console.log(fib(20));
"#), vec!["0", "1", "55", "6765"]);
}

#[test]
fn memoized_fibonacci() {
    assert_eq!(run_js(r#"
const memo = {};
function fib(n) {
    if (n in memo) return memo[n];
    if (n <= 1) return n;
    return memo[n] = fib(n-1) + fib(n-2);
}
console.log(fib(30));
console.log(fib(35));
"#), vec!["832040", "9227465"]);
}

#[test]
fn coin_change_greedy_check() {
    assert_eq!(run_js(r#"
function coinChange(coins, amount) {
    const dp = Array(amount + 1).fill(Infinity);
    dp[0] = 0;
    for (let i = 1; i <= amount; i++) {
        for (const coin of coins) {
            if (coin <= i) dp[i] = Math.min(dp[i], dp[i - coin] + 1);
        }
    }
    return dp[amount] === Infinity ? -1 : dp[amount];
}
console.log(coinChange([1, 5, 10, 25], 36));
console.log(coinChange([2], 3));
"#), vec!["3", "-1"]);
}

#[test]
fn longest_increasing_subsequence() {
    assert_eq!(run_js(r#"
function lis(arr) {
    const dp = Array(arr.length).fill(1);
    for (let i = 1; i < arr.length; i++) {
        for (let j = 0; j < i; j++) {
            if (arr[j] < arr[i]) dp[i] = Math.max(dp[i], dp[j] + 1);
        }
    }
    return Math.max(...dp);
}
console.log(lis([10, 9, 2, 5, 3, 7, 101, 18]));
console.log(lis([0, 1, 0, 3, 2, 3]));
"#), vec!["4", "4"]);
}

#[test]
fn longest_common_subsequence() {
    assert_eq!(run_js(r#"
function lcs(a, b) {
    const m = a.length, n = b.length;
    const dp = Array.from({length: m+1}, () => Array(n+1).fill(0));
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            if (a[i-1] === b[j-1]) dp[i][j] = dp[i-1][j-1] + 1;
            else dp[i][j] = Math.max(dp[i-1][j], dp[i][j-1]);
        }
    }
    return dp[m][n];
}
console.log(lcs("abcde", "ace"));
console.log(lcs("abc", "abc"));
console.log(lcs("abc", "def"));
"#), vec!["3", "3", "0"]);
}

#[test]
fn knapsack_01() {
    assert_eq!(run_js(r#"
function knapsack(weights, values, capacity) {
    const n = weights.length;
    const dp = Array.from({length: n+1}, () => Array(capacity+1).fill(0));
    for (let i = 1; i <= n; i++) {
        for (let w = 0; w <= capacity; w++) {
            dp[i][w] = dp[i-1][w];
            if (weights[i-1] <= w) {
                dp[i][w] = Math.max(dp[i][w], dp[i-1][w-weights[i-1]] + values[i-1]);
            }
        }
    }
    return dp[n][capacity];
}
console.log(knapsack([2, 3, 4, 5], [3, 4, 5, 6], 5));
"#), vec!["7"]);
}

#[test]
fn edit_distance() {
    assert_eq!(run_js(r#"
function editDistance(a, b) {
    const m = a.length, n = b.length;
    const dp = Array.from({length: m+1}, (_, i) => Array.from({length: n+1}, (_, j) => i || j));
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            if (a[i-1] === b[j-1]) dp[i][j] = dp[i-1][j-1];
            else dp[i][j] = 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
        }
    }
    return dp[m][n];
}
console.log(editDistance("horse", "ros"));
console.log(editDistance("intention", "execution"));
console.log(editDistance("", "abc"));
"#), vec!["3", "5", "3"]);
}

#[test]
fn max_profit_stock() {
    assert_eq!(run_js(r#"
function maxProfit(prices) {
    let min = Infinity, profit = 0;
    for (const p of prices) {
        min = Math.min(min, p);
        profit = Math.max(profit, p - min);
    }
    return profit;
}
console.log(maxProfit([7, 1, 5, 3, 6, 4]));
console.log(maxProfit([7, 6, 4, 3, 1]));
"#), vec!["5", "0"]);
}

#[test]
fn triangular_number_paths() {
    assert_eq!(run_js(r#"
function minTrianglePath(triangle) {
    const dp = [...triangle[triangle.length - 1]];
    for (let i = triangle.length - 2; i >= 0; i--) {
        for (let j = 0; j <= i; j++) {
            dp[j] = triangle[i][j] + Math.min(dp[j], dp[j+1]);
        }
    }
    return dp[0];
}
const t = [[2], [3, 4], [6, 5, 7], [4, 1, 8, 3]];
console.log(minTrianglePath(t));
"#), vec!["11"]);
}

#[test]
fn word_break_dp() {
    assert_eq!(run_js(r#"
function wordBreak(s, wordDict) {
    const set = new Set(wordDict);
    const dp = Array(s.length + 1).fill(false);
    dp[0] = true;
    for (let i = 1; i <= s.length; i++) {
        for (let j = 0; j < i; j++) {
            if (dp[j] && set.has(s.slice(j, i))) { dp[i] = true; break; }
        }
    }
    return dp[s.length];
}
console.log(wordBreak("leetcode", ["leet", "code"]));
console.log(wordBreak("applepenapple", ["apple", "pen"]));
console.log(wordBreak("catsandog", ["cats", "dog", "sand", "and", "cat"]));
"#), vec!["true", "true", "false"]);
}

#[test]
fn house_robber() {
    assert_eq!(run_js(r#"
function rob(nums) {
    let prev2 = 0, prev1 = 0;
    for (const n of nums) {
        const curr = Math.max(prev1, prev2 + n);
        prev2 = prev1;
        prev1 = curr;
    }
    return prev1;
}
console.log(rob([1, 2, 3, 1]));
console.log(rob([2, 7, 9, 3, 1]));
"#), vec!["4", "12"]);
}

#[test]
fn jump_game() {
    assert_eq!(run_js(r#"
function canJump(nums) {
    let maxReach = 0;
    for (let i = 0; i < nums.length; i++) {
        if (i > maxReach) return false;
        maxReach = Math.max(maxReach, i + nums[i]);
    }
    return true;
}
console.log(canJump([2, 3, 1, 1, 4]));
console.log(canJump([3, 2, 1, 0, 4]));
"#), vec!["true", "false"]);
}

#[test]
fn count_combinations() {
    assert_eq!(run_js(r#"
function combinations(n, k) {
    if (k === 0 || k === n) return 1;
    const dp = Array.from({length: n+1}, (_, i) => Array(i+1).fill(1));
    for (let i = 2; i <= n; i++) {
        for (let j = 1; j < i; j++) {
            dp[i][j] = dp[i-1][j-1] + dp[i-1][j];
        }
    }
    return dp[n][k];
}
console.log(combinations(4, 2));
console.log(combinations(5, 0));
console.log(combinations(6, 3));
"#), vec!["6", "1", "20"]);
}

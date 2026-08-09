/// Array algorithm patterns — sorting, searching, transforming
use super::helpers::run_js;

#[test]
fn binary_search() {
    assert_eq!(
        run_js(
            r#"
function binarySearch(arr, target) {
    let lo = 0, hi = arr.length - 1;
    while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (arr[mid] === target) return mid;
        else if (arr[mid] < target) lo = mid + 1;
        else hi = mid - 1;
    }
    return -1;
}
const arr = [1, 3, 5, 7, 9, 11, 13, 15];
console.log(binarySearch(arr, 7));
console.log(binarySearch(arr, 1));
console.log(binarySearch(arr, 15));
console.log(binarySearch(arr, 4));
"#
        ),
        vec!["3", "0", "7", "-1"]
    );
}

#[test]
fn binary_search_empty_array_returns_minus_one() {
    assert_eq!(
        run_js(
            r#"
function binarySearch(arr, target) {
    let lo = 0, hi = arr.length - 1;
    while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (arr[mid] === target) return mid;
        else if (arr[mid] < target) lo = mid + 1;
        else hi = mid - 1;
    }
    return -1;
}
console.log(binarySearch([], 10));
"#
        ),
        vec!["-1"]
    );
}

#[test]
fn quicksort() {
    assert_eq!(
        run_js(
            r#"
function quicksort(arr) {
    if (arr.length <= 1) return arr;
    const pivot = arr[arr.length >> 1];
    const left = arr.filter(x => x < pivot);
    const mid = arr.filter(x => x === pivot);
    const right = arr.filter(x => x > pivot);
    return [...quicksort(left), ...mid, ...quicksort(right)];
}
console.log(quicksort([3, 1, 4, 1, 5, 9, 2, 6]).join(","));
"#
        ),
        vec!["1,1,2,3,4,5,6,9"]
    );
}

#[test]
fn merge_sort() {
    assert_eq!(
        run_js(
            r#"
function merge(a, b) {
    const result = [];
    let i = 0, j = 0;
    while (i < a.length && j < b.length) {
        result.push(a[i] <= b[j] ? a[i++] : b[j++]);
    }
    return [...result, ...a.slice(i), ...b.slice(j)];
}
function mergeSort(arr) {
    if (arr.length <= 1) return arr;
    const mid = arr.length >> 1;
    return merge(mergeSort(arr.slice(0, mid)), mergeSort(arr.slice(mid)));
}
console.log(mergeSort([5, 3, 8, 1, 9, 2]).join(","));
"#
        ),
        vec!["1,2,3,5,8,9"]
    );
}

#[test]
fn two_sum() {
    assert_eq!(
        run_js(
            r#"
function twoSum(nums, target) {
    const map = new Map();
    for (let i = 0; i < nums.length; i++) {
        const complement = target - nums[i];
        if (map.has(complement)) return [map.get(complement), i];
        map.set(nums[i], i);
    }
    return [];
}
const [a, b] = twoSum([2, 7, 11, 15], 9);
console.log(a);
console.log(b);
console.log(twoSum([3, 2, 4], 6).join(","));
"#
        ),
        vec!["0", "1", "1,2"]
    );
}

#[test]
fn sliding_window_max() {
    assert_eq!(
        run_js(
            r#"
function maxSubarraySum(arr, k) {
    let sum = arr.slice(0, k).reduce((a, b) => a + b, 0);
    let max = sum;
    for (let i = k; i < arr.length; i++) {
        sum += arr[i] - arr[i - k];
        max = Math.max(max, sum);
    }
    return max;
}
console.log(maxSubarraySum([1, 3, -1, -3, 5, 3, 6, 7], 3));
"#
        ),
        vec!["16"]
    );
}

#[test]
fn kadanes_algorithm() {
    assert_eq!(
        run_js(
            r#"
function maxSubarray(arr) {
    let maxSum = arr[0], current = arr[0];
    for (let i = 1; i < arr.length; i++) {
        current = Math.max(arr[i], current + arr[i]);
        maxSum = Math.max(maxSum, current);
    }
    return maxSum;
}
console.log(maxSubarray([-2, 1, -3, 4, -1, 2, 1, -5, 4]));
console.log(maxSubarray([-1, -2, -3]));
"#
        ),
        vec!["6", "-1"]
    );
}

#[test]
fn fisher_yates_shuffle_length() {
    assert_eq!(
        run_js(
            r#"
function shuffle(arr) {
    const a = [...arr];
    for (let i = a.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
}
const orig = [1, 2, 3, 4, 5];
const shuffled = shuffle(orig);
console.log(shuffled.length);
console.log(shuffled.sort((a,b)=>a-b).join(","));
"#
        ),
        vec!["5", "1,2,3,4,5"]
    );
}

#[test]
fn rotate_array() {
    assert_eq!(
        run_js(
            r#"
function rotateLeft(arr, k) {
    const n = arr.length;
    k = k % n;
    return [...arr.slice(k), ...arr.slice(0, k)];
}
console.log(rotateLeft([1, 2, 3, 4, 5], 2).join(","));
console.log(rotateLeft([1, 2, 3], 0).join(","));
console.log(rotateLeft([1, 2, 3, 4], 4).join(","));
"#
        ),
        vec!["3,4,5,1,2", "1,2,3", "1,2,3,4"]
    );
}

#[test]
fn rotate_array_with_negative_shift_uses_slice_math() {
    assert_eq!(
        run_js(
            r#"
function rotateLeft(arr, k) {
    const n = arr.length;
    k = k % n;
    return [...arr.slice(k), ...arr.slice(0, k)];
}
console.log(rotateLeft([1, 2, 3], -1).join(","));
"#
        ),
        vec!["3,1,2"]
    );
}

#[test]
fn longest_common_prefix() {
    assert_eq!(
        run_js(
            r#"
function longestCommonPrefix(strs) {
    if (!strs.length) return "";
    return strs.reduce((prefix, str) => {
        while (!str.startsWith(prefix)) prefix = prefix.slice(0, -1);
        return prefix;
    });
}
console.log(longestCommonPrefix(["flower", "flow", "flight"]));
console.log(longestCommonPrefix(["dog", "racecar", "car"]));
console.log(longestCommonPrefix(["abc", "abcd", "ab"]));
"#
        ),
        vec!["fl", "", "ab"]
    );
}

#[test]
fn matrix_transpose() {
    assert_eq!(
        run_js(
            r#"
function transpose(matrix) {
    return matrix[0].map((_, i) => matrix.map(row => row[i]));
}
const m = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
const t = transpose(m);
console.log(t[0].join(","));
console.log(t[1].join(","));
console.log(t[2].join(","));
"#
        ),
        vec!["1,4,7", "2,5,8", "3,6,9"]
    );
}

#[test]
fn run_length_encoding() {
    assert_eq!(
        run_js(
            r#"
function rle(str) {
    return str.replace(/(.)\1*/g, (m, c) => m.length > 1 ? m.length + c : c);
}
console.log(rle("aabbbccddddee"));
console.log(rle("abc"));
console.log(rle("aaaa"));
"#
        ),
        vec!["2a3b2c4d2e", "abc", "4a"]
    );
}

#[test]
fn anagram_check() {
    assert_eq!(
        run_js(
            r#"
function isAnagram(a, b) {
    if (a.length !== b.length) return false;
    const sort = s => s.split("").sort().join("");
    return sort(a) === sort(b);
}
console.log(isAnagram("listen", "silent"));
console.log(isAnagram("hello", "world"));
console.log(isAnagram("anagram", "nagaram"));
"#
        ),
        vec!["true", "false", "true"]
    );
}

#[test]
fn frequency_map() {
    assert_eq!(
        run_js(
            r#"
function topK(arr, k) {
    const freq = new Map();
    for (const x of arr) freq.set(x, (freq.get(x) ?? 0) + 1);
    return [...freq.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, k)
        .map(([val]) => val);
}
const result = topK([1, 1, 1, 2, 2, 3], 2);
console.log(result[0]);
console.log(result[1]);
"#
        ),
        vec!["1", "2"]
    );
}

#[test]
fn zip_unzip() {
    assert_eq!(
        run_js(
            r#"
const zip = (...arrs) => arrs[0].map((_, i) => arrs.map(a => a[i]));
const unzip = arrs => arrs[0].map((_, i) => arrs.map(a => a[i]));
const zipped = zip([1, 2, 3], ["a", "b", "c"]);
console.log(zipped[0].join(","));
console.log(zipped[2].join(","));
"#
        ),
        vec!["1,a", "3,c"]
    );
}

#[test]
fn iterative_flat_deep_arrays() {
    assert_eq!(
        run_js(
            r#"
function flatDeep(arr) {
    const stack = [...arr];
    const res = [];
    while (stack.length) {
        const next = stack.pop();
        if (Array.isArray(next)) stack.push(...next);
        else res.push(next);
    }
    return res.reverse();
}
console.log(flatDeep([1, [2, [3, [4]]]]).join(","));
"#
        ),
        vec!["1,2,3,4"]
    );
}

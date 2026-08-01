use crate::helpers::run_prints;

#[test]
fn test_recursion_factorial_basic() {
    let out = run_prints(
        r#"
        fun fact(n: Int): Int = if (n <= 1) 1 else n * fact(n - 1)
        fun main() {
            println(fact(5))
        }
    "#,
    );
    assert_eq!(out, &["120"]);
}

#[test]
fn test_recursion_sum_down() {
    let out = run_prints(
        r#"
        fun sum(n: Int): Int = if (n <= 0) 0 else n + sum(n - 1)
        fun main() {
            println(sum(4))
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_recursion_power() {
    let out = run_prints(
        r#"
        fun pow(base: Int, exp: Int): Int = if (exp == 0) 1 else base * pow(base, exp - 1)
        fun main() {
            println(pow(2, 4))
        }
    "#,
    );
    assert_eq!(out, &["16"]);
}

#[test]
fn test_recursion_fibonacci() {
    let out = run_prints(
        r#"
        fun fib(n: Int): Int = if (n <= 1) n else fib(n - 1) + fib(n - 2)
        fun main() {
            println(fib(6))
        }
    "#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_recursion_mutual_even_odd() {
    let out = run_prints(
        r#"
        fun even(n: Int): Boolean = if (n == 0) true else odd(n - 1)
        fun odd(n: Int): Boolean = if (n == 0) false else even(n - 1)
        fun main() {
            println(even(10))
            println(odd(10))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_recursion_string_builder() {
    let out = run_prints(
        r#"
        fun repeatChar(ch: Char, count: Int): String {
            return if (count <= 0) "" else ch + repeatChar(ch, count - 1)
        }
        fun main() {
            println(repeatChar('x', 3))
        }
    "#,
    );
    assert_eq!(out, &["xxx"]);
}

#[test]
fn test_recursion_list_count() {
    let out = run_prints(
        r#"
        fun count(items: List<Int>): Int = if (items.isEmpty()) 0 else 1 + count(items.drop(1))
        fun main() {
            println(count(listOf(1, 2, 3)))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_recursion_head_sum() {
    let out = run_prints(
        r#"
        fun headSum(values: List<Int>): Int = if (values.isEmpty()) 0 else values[0] + headSum(values.drop(1))
        fun main() {
            println(headSum(listOf(2, 4, 6)))
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_recursion_find_max() {
    let out = run_prints(
        r#"
        fun maxOf(values: List<Int>): Int {
            if (values.size == 1) return values[0]
            val tail = values.drop(1)
            val candidate = maxOf(tail)
            return if (values[0] > candidate) values[0] else candidate
        }
        fun main() {
            println(maxOf(listOf(3, 1, 9, 2)))
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_recursion_reverse_string() {
    let out = run_prints(
        r#"
        fun rev(s: String): String = if (s.isEmpty()) "" else rev(s.substring(1)) + s[0]
        fun main() {
            println(rev("abc"))
        }
    "#,
    );
    assert_eq!(out, &["cba"]);
}

#[test]
fn test_recursion_nested_recursion() {
    let out = run_prints(
        r#"
        fun nested(n: Int): Int = if (n == 0) 1 else n * nested(n - 1) + nested(n - 2)
        fun main() {
            println(nested(4))
        }
    "#,
    );
    assert_eq!(out, &["34"]);
}

#[test]
fn test_recursion_while_to_tail_style() {
    let out = run_prints(
        r#"
        fun countdown(n: Int): Int {
            return if (n == 0) 0 else 1 + countdown(n - 1)
        }
        fun main() {
            println(countdown(0))
            println(countdown(2))
        }
    "#,
    );
    assert_eq!(out, &["0", "2"]);
}

#[test]
fn test_recursion_path_sum() {
    let out = run_prints(
        r#"
        fun path(n: Int, acc: Int): Int = if (n <= 0) acc else path(n - 1, acc + n)
        fun main() {
            println(path(3, 0))
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_recursion_gcd_recursive() {
    let out = run_prints(
        r#"
        fun gcd(a: Int, b: Int): Int = if (b == 0) a else gcd(b, a % b)
        fun main() {
            println(gcd(20, 12))
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_recursion_nested_if_even_odd_sum() {
    let out = run_prints(
        r#"
        fun altSum(n: Int): Int = if (n <= 0) 0 else if (n % 2 == 0) altSum(n - 1) - n else altSum(n - 1) + n
        fun main() {
            println(altSum(4))
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_recursion_tree_depth() {
    let out = run_prints(
        r#"
        class Node(val left: Node?, val right: Node?, val value: Int)
        fun depth(node: Node?): Int = if (node == null) 0 else 1 + maxOf(depth(node.left), depth(node.right))
        fun maxOf(a: Int, b: Int): Int = if (a > b) a else b
        fun main() {
            val t = Node(Node(null, null, 2), null, 1)
            println(depth(t))
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_recursion_binary_length() {
    let out = run_prints(
        r#"
        fun binLen(n: Int): Int = if (n < 2) 1 else 1 + binLen(n / 2)
        fun main() {
            println(binLen(8))
            println(binLen(1))
        }
    "#,
    );
    assert_eq!(out, &["4", "1"]);
}

#[test]
fn test_recursion_string_split_depth() {
    let out = run_prints(
        r#"
        fun words(s: String): Int = if (s.isEmpty()) 0 else 1 + words(s.substring(1).trimStart())
        fun main() {
            println(words("x"))
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_recursion_repeat_list() {
    let out = run_prints(
        r#"
        fun repeat(value: String, count: Int): String = if (count <= 0) "" else value + repeat(value, count - 1)
        fun main() {
            println(repeat("a", 2))
        }
    "#,
    );
    assert_eq!(out, &["aa"]);
}

#[test]
fn test_recursion_palindrome_check() {
    let out = run_prints(
        r#"
        fun isPal(s: String): Boolean {
            if (s.length <= 1) return true
            if (s.first() != s.last()) return false
            return isPal(s.substring(1, s.length - 1))
        }
        fun main() {
            println(isPal("racecar"))
            println(isPal("kotlin"))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_recursion_sum_nested_arrays() {
    let out = run_prints(
        r#"
        fun nestedSum(items: List<List<Int>>, row: Int = 0): Int {
            return if (row >= items.size) 0 else items[row].sum() + nestedSum(items, row + 1)
        }
        fun main() {
            println(nestedSum(listOf(listOf(1, 2), listOf(3, 4))))
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_recursion_count_char() {
    let out = run_prints(
        r#"
        fun countChar(s: String, target: Char): Int {
            if (s.isEmpty()) return 0
            val head = if (s[0] == target) 1 else 0
            return head + countChar(s.substring(1), target)
        }
        fun main() {
            println(countChar("abca", 'a'))
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_recursion_tail_ish_sum() {
    let out = run_prints(
        r#"
        fun sumTail(n: Int, acc: Int = 0): Int {
            return if (n <= 0) acc else sumTail(n - 1, acc + n)
        }
        fun main() {
            println(sumTail(4))
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_recursion_nested_while_simulation() {
    let out = run_prints(
        r#"
        fun dropCount(values: List<Int>): Int {
            return if (values.isEmpty()) 0 else 1 + dropCount(values.drop(1).drop(1))
        }
        fun main() {
            println(dropCount(listOf(1, 2, 3, 4)))
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_recursion_triangle_numbers() {
    let out = run_prints(
        r#"
        fun triangle(n: Int): Int = if (n <= 0) 0 else n + triangle(n - 1)
        fun main() {
            println(triangle(5))
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}

#[test]
fn test_recursion_string_array_joiner() {
    let out = run_prints(
        r#"
        fun join(values: List<String>): String = if (values.isEmpty()) "" else values[0] + if (values.size == 1) "" else "," + join(values.drop(1))
        fun main() {
            println(join(listOf("a", "b", "c")))
        }
    "#,
    );
    assert_eq!(out, &["a,b,c"]);
}

#[test]
fn test_recursion_countdown_string() {
    let out = run_prints(
        r#"
        fun dump(n: Int): String = if (n <= 0) "0" else n.toString() + "," + dump(n - 1)
        fun main() {
            println(dump(3))
        }
    "#,
    );
    assert_eq!(out, &["3,2,1,0"]);
}

#[test]
fn test_recursion_depth_guard() {
    let out = run_prints(
        r#"
        fun depth(v: Int): Int {
            if (v <= 0) return 0
            return if (v == 1) 1 else 1 + depth(v - 1)
        }
        fun main() {
            println(depth(1))
            println(depth(4))
        }
    "#,
    );
    assert_eq!(out, &["1", "4"]);
}

#[test]
fn test_recursion_sum_map_values() {
    let out = run_prints(
        r#"
        fun sumMap(values: Map<String, Int>): Int {
            if (values.isEmpty()) return 0
            val first = values.entries.first()
            return first.value + sumMap(values - first.key)
        }
        fun main() {
            println(sumMap(mapOf("a" to 1, "b" to 2)))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_recursion_digit_sum() {
    let out = run_prints(
        r#"
        fun digitSum(v: Int): Int = if (v == 0) 0 else (v % 10) + digitSum(v / 10)
        fun main() {
            println(digitSum(1234))
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_recursion_range_step() {
    let out = run_prints(
        r#"
        fun countRange(start: Int, end: Int): Int = if (start > end) 0 else 1 + countRange(start + 1, end)
        fun main() {
            println(countRange(1, 3))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_recursion_nested_map() {
    let out = run_prints(
        r#"
        fun flatten(items: List<List<Int>>, row: Int = 0): List<Int> {
            if (row >= items.size) return listOf()
            return items[row] + flatten(items, row + 1)
        }
        fun main() {
            val values = flatten(listOf(listOf(1), listOf(2, 3)))
            println(values.joinToString("."))
        }
    "#,
    );
    assert_eq!(out, &["1.2.3"]);
}

#[test]
fn test_recursion_bounded_repeat() {
    let out = run_prints(
        r#"
        fun repeatText(value: String, count: Int): String = if (count <= 0) "" else value + " " + repeatText(value, count - 1)
        fun main() {
            println(repeatText("x", 2))
        }
    "#,
    );
    assert_eq!(out, &["x x "]);
}

#[test]
fn test_recursion_binary_grow() {
    let out = run_prints(
        r#"
        fun seq(n: Int): Int = if (n <= 1) 1 else seq(n - 1) + seq(n - 2)
        fun main() {
            println(seq(7))
        }
    "#,
    );
    assert_eq!(out, &["13"]);
}

#[test]
fn test_recursion_nested_tail_guard() {
    let out = run_prints(
        r#"
        fun nested(n: Int): Int {
            if (n == 0) return 1
            if (n == 1) return 1
            return nested(n - 1) + nested(n - 2)
        }
        fun main() {
            println(nested(6))
        }
    "#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_recursion_empty_input_handling() {
    let out = run_prints(
        r#"
        fun total(values: List<Int>): Int {
            if (values.isEmpty()) return 0
            return values.first() + total(values.drop(1))
        }
        fun main() {
            println(total(listOf<Int>()))
            println(total(listOf(9)))
        }
    "#,
    );
    assert_eq!(out, &["0", "9"]);
}

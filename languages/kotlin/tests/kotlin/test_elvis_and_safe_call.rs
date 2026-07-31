kotlin_run_test!(
    test_safe_call_null,
    r#"fun main() { val x: String? = null; println(x?.length) }"#,
    &["null"]
);

kotlin_run_test!(
    test_safe_call_non_null,
    r#"fun main() { val x: String? = "abc"; println(x?.length) }"#,
    &["3"]
);

kotlin_run_test!(
    test_safe_call_chain,
    r#"class N { val v: Int? = 2 }
fun main() { val n: N? = N(); println(n?.v?.plus(1)) }"#,
    &["3"]
);

kotlin_run_test!(
    test_safe_call_break_in_chain,
    r#"class N { val v: Int? = null }
fun main() { val n: N? = N(); println(n?.v?.plus(1)) }"#,
    &["null"]
);

kotlin_run_test!(
    test_elvis_on_null,
    r#"fun main() { val x: String? = null; println(x ?: "none") }"#,
    &["none"]
);

kotlin_run_test!(
    test_elvis_on_non_null,
    r#"fun main() { val x: String? = "go"; println(x ?: "none") }"#,
    &["go"]
);

kotlin_run_test!(
    test_elvis_with_function,
    r#"fun fallback(): String = "fb"
fun main() { val x: String? = null; println((x ?: fallback()).length) }"#,
    &["2"]
);

kotlin_run_test!(
    test_elvis_default_callable,
    r#"fun main() {
        fun choose(v: String?): String = v ?: run { println("fallback") ; "x" }
        println(choose(null))
    }"#,
    &["fallback", "x"]
);

kotlin_run_test!(
    test_elvis_return,
    r#"fun title(v: String?): String {
        return v ?: return "none"
    }
fun main() { println(title(null)); println(title("x")) }"#,
    &["none", "x"]
);

kotlin_run_test!(
    test_safe_cast,
    r#"fun main() { val x: Any? = "abc"; val y = x as? Int; println(y) }"#,
    &["null"]
);

kotlin_run_test!(
    test_safe_cast_success,
    r#"fun main() { val x: Any? = 5; val y = x as? Int; println(y) }"#,
    &["5"]
);

kotlin_run_test!(
    test_non_null_assertion_success,
    r#"fun main() { val x: String? = "ok"; println(x!!); }"#,
    &["ok"]
);

kotlin_run_test!(
    test_non_null_assertion_failure,
    r#"fun main() { try { val x: String? = null; println(x!!) } catch (e: Exception) { println("err") } }"#,
    &["err"]
);

kotlin_run_test!(
    test_elvis_with_expression,
    r#"fun main() {
        val x: Int? = null
        val y = x?.plus(1) ?: 9
        println(y)
    }"#,
    &["9"]
);

kotlin_run_test!(
    test_elvis_chain_multiple,
    r#"fun first(): String? = null
fun second(): String? = "ok"
fun main() { println(first() ?: second() ?: "z") }"#,
    &["ok"]
);

kotlin_run_test!(
    test_elvis_on_nullable_list,
    r#"fun main() { val x: List<Int>? = null; println((x?.size ?: 0) + 1) }"#,
    &["1"]
);

kotlin_run_test!(
    test_safe_call_in_when,
    r#"fun main() {
        val x: String? = null
        when (x?.length ?: 0) {
            0 -> println("z")
            else -> println("n")
        }
    }"#,
    &["z"]
);

kotlin_run_test!(
    test_safe_call_in_loops,
    r#"fun main() {
        val xs: List<String?> = listOf(null, "a", null, "bb")
        var c = 0
        for (s in xs) { c += s?.length ?: 0 }
        println(c)
    }"#,
    &["3"]
);

kotlin_run_test!(
    test_elvis_list_default,
    r#"fun main() { val xs: List<Int>? = null; val out = xs ?: listOf(1,2); println(out.size) }"#,
    &["2"]
);

kotlin_run_test!(
    test_safe_call_in_map_lookup,
    r#"fun main() {
        val m: Map<String, String>? = null
        println(m?.get("a") ?: "na")
    }"#,
    &["na"]
);

kotlin_run_test!(
    test_null_guard_with_let,
    r#"fun main() {
        val x: String? = "ok"
        x?.let { println("v" + it) } ?: println("missing")
    }"#,
    &["vok"]
);

kotlin_run_test!(
    test_null_guard_without_let,
    r#"fun main() {
        val x: String? = null
        x?.let { println("v" + it) } ?: println("missing")
    }"#,
    &["missing"]
);

kotlin_run_test!(
    test_chained_safe_calls,
    r#"class A { fun child() = B() }
class B { fun name(): String = "b" }
fun main() { val a: A? = A(); println(a?.child()?.name()) }"#,
    &["b"]
);

kotlin_run_test!(
    test_safe_call_array_index,
    r#"fun main() {
        val a: IntArray? = null
        println(a?.get(0) ?: -1)
    }"#,
    &["-1"]
);

kotlin_run_test!(
    test_safe_call_array_present,
    r#"fun main() {
        val a: IntArray? = intArrayOf(4,5)
        println(a?.get(1) ?: -1)
    }"#,
    &["5"]
);

kotlin_run_test!(
    test_nullable_return_chain,
    r#"fun find(v: Boolean): String? = if (v) "yes" else null
fun main() { val out = find(false) ?: find(true) ?: "none"; println(out) }"#,
    &["yes"]
);

kotlin_run_test!(
    test_elvis_with_nullable_receiver,
    r#"fun main() {
        val s: String? = null
        println(s?.uppercase() ?: "NONE")
    }"#,
    &["NONE"]
);

kotlin_run_test!(
    test_elvis_with_method_call,
    r#"fun main() {
        val s: String? = "ok"
        println(s?.uppercase()?.substring(0, 1) ?: "missing")
    }"#,
    &["O"]
);

kotlin_run_test!(
    test_elvis_numeric_default,
    r#"fun main() {
        val x: Int? = null
        println((x ?: 2) * 3)
    }"#,
    &["6"]
);

kotlin_run_test!(
    test_safe_call_list_first,
    r#"fun main() {
        val xs: List<String>? = listOf("a", "b")
        println(xs?.firstOrNull() ?: "none")
    }"#,
    &["a"]
);

kotlin_run_test!(
    test_safe_call_list_empty,
    r#"fun main() {
        val xs: List<String>? = listOf()
        println(xs?.firstOrNull() ?: "none")
    }"#,
    &["none"]
);

kotlin_run_test!(
    test_safe_call_map_values,
    r#"fun main() {
        val x: Map<String, Int>? = mapOf("a" to 2)
        println(x?.get("a") ?: -1)
    }"#,
    &["2"]
);

kotlin_run_test!(
    test_safe_call_nested_map,
    r#"fun main() {
        val x: Map<String, Map<String, Int>?>? = mapOf("a" to mapOf("b" to 3))
        println(x?.get("a")?.get("b") ?: -1)
    }"#,
    &["3"]
);

kotlin_run_test!(
    test_safe_call_nested_map_missing,
    r#"fun main() {
        val x: Map<String, Map<String, Int>?>? = mapOf("a" to null)
        println(x?.get("a")?.get("b") ?: -1)
    }"#,
    &["-1"]
);

kotlin_run_test!(
    test_safe_call_in_string_concat,
    r#"fun main() {
        val s: String? = null
        println((s ?: "") + "x")
    }"#,
    &["x"]
);

kotlin_run_test!(
    test_notnull_cast_and_safe,
    r#"fun main() {
        val x: Any? = "ok"
        println((x as? String)?.length ?: 0)
    }"#,
    &["2"]
);

kotlin_run_test!(
    test_nonnull_throwing_with_elvis,
    r#"fun main() {
        val x: String? = ""
        println(x!!.ifEmpty { "empty" })
    }"#,
    &["empty"]
);

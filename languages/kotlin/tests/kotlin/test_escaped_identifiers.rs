kotlin_run_test!(
    test_keyword_as_variable,
    r#"fun main() { val `class` = 7; println(`class`) }"#,
    &["7"]
);

kotlin_run_test!(
    test_keyword_as_function,
    r#"fun `when`(): Int = 5
fun main() { println(`when`()) }"#,
    &["5"]
);

kotlin_run_test!(
    test_space_in_identifier,
    r#"fun `add one`(x: Int): Int = x + 1
fun main() { println(`add one`(2)) }"#,
    &["3"]
);

kotlin_run_test!(
    test_dash_in_identifier,
    r#"fun `value-sum`(a: Int, b: Int): Int = a + b
fun main() { println(`value-sum`(2, 3)) }"#,
    &["5"]
);

kotlin_run_test!(
    test_underscored_backtick,
    r#"fun main() { val `a b_c` = 4; println(`a b_c`) }"#,
    &["4"]
);

kotlin_run_test!(
    test_class_with_space_name,
    r#"class `My Class` { val value = 1 }
fun main() { println(`My Class`().value) }"#,
    &["1"]
);

kotlin_run_test!(
    test_property_with_space,
    r#"class Holder { val `space key` = 9 }
fun main() { val h = Holder(); println(h.`space key`) }"#,
    &["9"]
);

kotlin_run_test!(
    test_function_with_symbols,
    r#"fun `f+g`(a: Int): Int = a + 10
fun main() { println(`f+g`(1)) }"#,
    &["11"]
);

kotlin_run_test!(
    test_function_with_operator_like_name,
    r#"fun `a b`(x: Int, y: Int): Int = x * y
fun main() { println(`a b`(3, 4)) }"#,
    &["12"]
);

kotlin_run_test!(
    test_data_class_with_backtick_field,
    r#"data class `Item Box`(val `item id`: Int)
fun main() { val i = `Item Box`(4); println(i.`item id`) }"#,
    &["4"]
);

kotlin_run_test!(
    test_method_call_with_backtick_name,
    r#"class Api { fun `compute next`(x: Int) = x + 2 }
fun main() { val a = Api(); println(a.`compute next`(5)) }"#,
    &["7"]
);

kotlin_run_test!(
    test_extension_on_string_backtick,
    r#"fun String.`dash`(): String = this + "!"
fun main() { println("ok".`dash`()) }"#,
    &["ok!"]
);

kotlin_run_test!(
    test_interface_with_backtick_member,
    r#"interface `I-Thing` { val `prop value`: Int }
class C: `I-Thing` { override val `prop value` = 11 }
fun main() { println(C().`prop value`) }"#,
    &["11"]
);

kotlin_run_test!(
    test_object_expression_backtick_property,
    r#"fun main() {
            val o = object { val `x y` = 3 }
            println(o.`x y`)
        }"#,
    &["3"]
);

kotlin_run_test!(
    test_backtick_constructor_parameter,
    r#"class Box(val `label text`: String)
fun main() { val b = Box("x"); println(b.`label text`) }"#,
    &["x"]
);

kotlin_run_test!(
    test_nested_backtick_class,
    r#"class Outer {
        inner class `Inner Type`(val `count value`: Int)
    }
    fun main() {
        println(Outer().`Inner Type`(7).`count value`)
    }"#,
    &["7"]
);

kotlin_run_test!(
    test_backtick_generic_type_alias,
    r#"typealias `String-ID` = String
fun main() { val id: `String-ID` = "k"; println(id) }"#,
    &["k"]
);

kotlin_run_test!(
    test_reserved_name_in_lambda,
    r#"fun main() {
        val `if` = { x: Int -> x + 1 }
        println(`if`(2))
    }"#,
    &["3"]
);

kotlin_run_test!(
    test_reserved_name_as_object,
    r#"fun main() {
        val `val` = 3
        val `var` = 4
        println(`val` + `var`)
    }"#,
    &["7"]
);

kotlin_run_test!(
    test_backtick_method_chain,
    r#"class Counter {
        fun `next value`(x: Int) = x + 1
        fun `next value`(x: Int, y: Int) = x + y
    }
    fun main() {
        val c = Counter()
        println(c.`next value`(3) + c.`next value`(1, 2))
    }"#,
    &["7"]
);

kotlin_run_test!(
    test_backtick_parameter_name,
    r#"fun combine(`first part`: Int, `second part`: Int) = `first part` + `second part`
fun main() { println(combine(4, 5)) }"#,
    &["9"]
);

kotlin_run_test!(
    test_backtick_field_setter,
    r#"class Holder { var `mutable field` = 1
            fun inc() { `mutable field` += 2 }
        }
    fun main() { val h = Holder(); h.inc(); println(h.`mutable field`) }"#,
    &["3"]
);

kotlin_run_test!(
    test_backtick_in_local_function,
    r#"fun main() {
        fun `local op`(x: Int, y: Int) = x * y
        println(`local op`(2, 6))
    }"#,
    &["12"]
);

kotlin_run_test!(
    test_backtick_in_object_function,
    r#"fun main() {
        val o = object {
            fun `compute`(x: Int): Int = x / 2
        }
        println(o.`compute`(8))
    }"#,
    &["4"]
);

kotlin_run_test!(
    test_backtick_top_level_constant,
    r#"val `global value` = 10
fun main() { println(`global value`) }"#,
    &["10"]
);

kotlin_run_test!(
    test_backtick_inheritance_override,
    r#"open class Base { open fun `do work`(): Int = 1 }
class Child: Base() { override fun `do work`() = 3 }
fun main() { println(Child().`do work`()) }"#,
    &["3"]
);

kotlin_run_test!(
    test_backtick_package_class_name_not_used,
    r#"class `X-Class` { fun value() = 9 }
fun main() { println(`X-Class`().value()) }"#,
    &["9"]
);

kotlin_run_test!(
    test_backtick_and_nested_calls,
    r#"class A { fun `outer`(b: String) = b }
fun main() { println(A().`outer`("go")) }"#,
    &["go"]
);

kotlin_run_test!(
    test_backtick_in_data_copy,
    r#"data class `Pair Data`(val `left value`: Int, val `right value`: Int)
fun main() {
    val p = `Pair Data`(1, 2)
    val q = p.copy(`right value` = 3)
    println(q.`left value` + q.`right value`)
}"#,
    &["4"]
);

kotlin_run_test!(
    test_backtick_in_destructuring,
    r#"data class `Node Pair`(val `a key`: Int, val `b key`: Int)
fun main() {
    val (`a key`, `b key`) = Pair(1, 2)
    println(`a key` + `b key`)
}"#,
    &["3"]
);

kotlin_run_test!(
    test_backtick_in_lambda_parameter,
    r#"fun main() {
        val fn = { `arg value`: Int -> `arg value` + 1 }
        println(fn(6))
    }"#,
    &["7"]
);

kotlin_run_test!(
    test_backtick_in_generic_type_name,
    r#"typealias `Alias Name` = Map<String, Int>
fun main() {
    val x: `Alias Name` = mapOf("a" to 1)
    println(x["a"])
}"#,
    &["1"]
);

kotlin_run_test!(
    test_backtick_in_when_subject,
    r#"fun main() {
        val `state value` = "ok"
        val out = when (`state value`) {
            "ok" -> "yes"
            else -> "no"
        }
        println(out)
    }"#,
    &["yes"]
);

kotlin_run_test!(
    test_backtick_in_type_parameter,
    r#"class `Holder Type`<T>(val `value data`: T)
fun main() {
    val h = `Holder Type`(5)
    println(h.`value data`)
}"#,
    &["5"]
);

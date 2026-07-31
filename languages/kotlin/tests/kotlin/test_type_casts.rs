use crate::helpers::run_prints;

#[test]
fn test_is_type_check() {
    let out = run_prints(r#"
        fun main() {
            val str = "hello"
            if (str is String) {
                println("is string")
            }
            if (str !is Int) {
                println("not int")
            }
        }
    "#);
    assert_eq!(out, &["is string", "not int"]);
}

#[test]
fn test_as_type_cast() {
    let out = run_prints(r#"
        fun main() {
            val obj = "kotlin language"
            val text = obj as String
            println(text)
        }
    "#);
    assert_eq!(out, &["kotlin language"]);
}

#[test]
fn test_is_check_with_boolean() {
    let out = run_prints(r#"
        fun main() {
            val flag = true
            if (flag is Boolean) {
                println("is boolean")
            }
        }
    "#);
    assert_eq!(out, &["is boolean"]);
}

#[test]
fn test_safe_cast_to_wrong_type() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = 100
            val casted = value as? String
            println(casted == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_safe_cast_success() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = "hello"
            val casted = value as? String
            println(casted!!)
        }
    "#);
    assert_eq!(out, &["hello"]);
}

#[test]
fn test_smart_cast_after_is() {
    let out = run_prints(r#"
        fun main() {
            val input: Any = 42
            if (input is Int) {
                val n: Int = input
                println(n + 1)
            } else {
                println(0)
            }
        }
    "#);
    assert_eq!(out, &["43"]);
}

#[test]
fn test_cast_between_class_hierarchy() {
    let out = run_prints(r#"
        open class Vehicle(val speed: Int)
        class Car(speed: Int) : Vehicle(speed)

        fun main() {
            val value: Vehicle = Car(120)
            val car = value as Car
            println(car.speed)
        }
    "#);
    assert_eq!(out, &["120"]);
}

#[test]
fn test_safe_cast_to_wrong_reference() {
    let out = run_prints(r#"
        open class Base
        class Child : Base

        fun main() {
            val base: Base = Base()
            val child = base as? Child
            println(child == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_not_is_false_when_matches() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = "text"
            if (value !is Int) {
                println("not_int")
            }
        }
    "#);
    assert_eq!(out, &["not_int"]);
}

#[test]
fn test_is_operator_true_branch() {
    let out = run_prints(r#"
        open class Node
        class Leaf : Node()

        fun main() {
            val node: Node = Leaf()
            if (node is Node) {
                println("is_node")
            }
        }
    "#);
    assert_eq!(out, &["is_node"]);
}

#[test]
fn test_as_cast_on_any_reference() {
    let out = run_prints(r#"
        fun main() {
            val any: Any = 77
            val num = any as Int
            println(num + 1)
        }
    "#);
    assert_eq!(out, &["78"]);
}

#[test]
fn test_safe_cast_with_null_source() {
    let out = run_prints(r#"
        fun main() {
            val source: Any? = null
            val casted = source as? String
            println(casted == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_smart_cast_after_if() {
    let out = run_prints(r#"
        fun main() {
            val item: Any = 55
            if (item is Int) {
                val value = item
                println(value * 2)
            }
        }
    "#);
    assert_eq!(out, &["110"]);
}

#[test]
fn test_nested_cast_in_expression() {
    let out = run_prints(r#"
        fun convert(input: Any): Int {
            return if (input is Int) {
                input as Int
            } else {
                0
            }
        }

        fun main() {
            println(convert(9))
            println(convert("bad"))
        }
    "#);
    assert_eq!(out, &["9", "0"]);
}

#[test]
fn test_cast_with_function_return() {
    let out = run_prints(r#"
        fun toText(value: Any): String {
            val casted = value as? String
            return casted ?: "fallback"
        }

        fun main() {
            println(toText("hello"))
            println(toText(2))
        }
    "#);
    assert_eq!(out, &["hello", "fallback"]);
}

#[test]
fn test_multiple_type_checks() {
    let out = run_prints(r#"
        fun check(value: Any) {
            if (value is String) {
                println("string")
            } else if (value !is Int) {
                println("not int")
            } else {
                println("int")
            }
        }

        fun main() {
            check("x")
            check(3)
            check(true)
        }
    "#);
    assert_eq!(out, &["string", "int", "not int"]);
}

#[test]
fn test_is_check_null() {
    let out = run_prints(r#"
fun main() { val value: Any? = null; println(value is Int) }
"#);
    assert_eq!(out, &["false"]);
}

#[test]
fn test_is_check_not_null() {
    let out = run_prints(r#"
fun main() { val value: Any? = "abc"; println(value is String) }
"#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_as_cast_from_any() {
    let out = run_prints(r#"
fun main() { val value: Any = 9; val casted = value as Int; println(casted + 10) }
"#);
    assert_eq!(out, &["19"]);
}

#[test]
fn test_safe_cast_returns_null() {
    let out = run_prints(r#"
fun main() { val value: Any = 9; println((value as? String) == null) }
"#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_safe_cast_success_in_if() {
    let out = run_prints(r#"
fun describe(value: Any): String { return if (value is Int) "int" else "not int" }; fun main() { println(describe(10)); println(describe("x")) }
"#);
    assert_eq!(out, &["int", "not int"]);
}

#[test]
fn test_nested_cast_chain() {
    let out = run_prints(r#"
open class A; class B : A(); fun castOrZero(v: A): Int { return if (v is B) 1 else 0 }; fun main() { println(castOrZero(B())); println(castOrZero(A())) }
"#);
    assert_eq!(out, &["1", "0"]);
}

#[test]
fn test_cast_with_optional_source() {
    let out = run_prints(r#"
fun toNumber(value: Any?): Int { return (value as? Int) ?: 0 }; fun main() { println(toNumber(null)); println(toNumber(4)) }
"#);
    assert_eq!(out, &["0", "4"]);
}

#[test]
fn test_as_in_typed_function() {
    let out = run_prints(r#"
fun getText(any: Any): String { return any as String }; fun main() { println(getText("ok")) }
"#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_is_not_check() {
    let out = run_prints(r#"
fun main() { val value: Any = 12; println(value !is String) }
"#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_boolean_is_cast() {
    let out = run_prints(r#"
fun main() { val value: Any = true; println(value is Boolean) }
"#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_boolean_as_cast_in_function() {
    let out = run_prints(r#"
fun isTrue(value: Any?): Boolean { return value as? Boolean ?: false }; fun main() { println(isTrue(true)); println(isTrue("n")) }
"#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_double_is_check() {
    let out = run_prints(r#"
fun main() { val value: Any = 1.5; println(value is Double) }
"#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_cascading_casts() {
    let out = run_prints(r#"
open class Base; class Child : Base(); fun main() { val any: Base = Child(); val child = any as Child; println(child is Child) }
"#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_safe_cast_of_nullables() {
    let out = run_prints(r#"
fun valueOrDefault(value: Any?): Int { val text = value as? Int; return text ?: 11 }; fun main() { println(valueOrDefault(null)); println(valueOrDefault(8)) }
"#);
    assert_eq!(out, &["11", "8"]);
}

#[test]
fn test_unsafe_cast_to_wrong_type_is_caught() {
    let out = run_prints(r#"
        fun main() {
            try {
                val value: Any = true
                val number = value as Int
                println(number)
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#);
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_as_nullable_target_with_null_source() {
    let out = run_prints(r#"
        fun main() {
            val value: Any? = null
            val casted: String? = value as String?
            println(casted == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_when_type_check_smart_casts() {
    let out = run_prints(r#"
        fun describe(value: Any): String {
            return when {
                value is String -> "string:" + value.length
                value is Int -> "int:" + (value + 1)
                value is Boolean -> "bool:" + (if (value) 1 else 0)
                else -> "other"
            }
        }

        fun main() {
            println(describe("kotlin"))
            println(describe(6))
            println(describe(false))
            println(describe(1.5))
        }
    "#);
    assert_eq!(out, &["string:6", "int:7", "bool:0", "other"]);
}

#[test]
fn test_sibling_downcast_failure_path() {
    let out = run_prints(r#"
        open class Shape
        class Circle : Shape()
        class Square : Shape()

        fun main() {
            val value: Shape = Square()
            try {
                val casted = value as Circle
                println(casted == null)
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#);
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_safe_cast_to_wrong_interface_is_none() {
    let out = run_prints(r#"
        interface Reader { fun read(): String }
        class FileReader : Reader {
            override fun read(): String = "ok"
        }

        fun main() {
            val value: Any = 99
            val reader = value as? Reader
            println(reader == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_and_guard_with_is_check() {
    let out = run_prints(r#"
        fun main() {
            val item: Any = "abc"
            if (item is String && item.isNotEmpty()) {
                println(item.length)
            } else {
                println(0)
            }
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_cast_then_safe_cast_chain() {
    let out = run_prints(r#"
fun main() {
    val source: Any? = "x"
            val direct = source as String
            val safe = (source as? String) ?: "fallback"
            val failed = source as? Int
            println(direct + ":" + safe)
    println(failed == null)
}
"#);
    assert_eq!(out, &["x:x", "true"]);
}

#[test]
fn test_nullable_is_nullable_type() {
    let out = run_prints(r#"
        fun main() {
            val value: String? = null
            println(value is String?)
            println(value is String)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_as_to_non_nullable_from_null_throws() {
    let out = run_prints(r#"
        fun main() {
            try {
                val value: String? = null
                val forced = value as String
                println("bad")
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#);
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_safe_cast_between_independent_interfaces() {
    let out = run_prints(r#"
        interface First { fun first(): Int }
        interface Second { fun second(): String }

        class Impl : First, Second {
            override fun first(): Int = 7
            override fun second(): String = "ok"
        }

        fun main() {
            val value: Any = Impl()
            val first = value as First
            val second = value as? Second
            if (second != null) {
                println(first.first().toString() + ":" + second.second())
            } else {
                println("missing")
            }
        }
    "#);
    assert_eq!(out, &["7:ok"]);
}

#[test]
fn test_safe_cast_to_missing_interface_is_null() {
    let out = run_prints(r#"
        interface Aware { fun marker(): String }
        interface Other { fun other(): Int }
        class Item : Aware { override fun marker(): String = "x" }

        fun main() {
            val value: Any = Item()
            val casted = value as? Other
            println(casted == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_is_number_type_checks() {
    let out = run_prints(r#"
        fun main() {
            val intValue: Any = 12
            val doubleValue: Any = 1.5
            println(intValue is Int)
            println(doubleValue is Int)
            println(doubleValue is Double)
        }
    "#);
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_as_array_to_incompatible_component_type_fails() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = arrayOf(1, 2, 3)
            try {
                val casted = value as Array<String>
                println(casted[0])
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#);
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_safe_cast_array_to_incompatible_component_type_is_null() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = arrayOf("a", "b")
            val casted = value as? Array<Int>
            println(casted == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_as_nullable_array_from_null_fails() {
    let out = run_prints(r#"
        fun main() {
            try {
                val value: Any? = null
                val casted = value as Array<Int>
                println(casted.size)
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#);
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_mutable_list_cast_to_readonly_list_and_back() {
    let out = run_prints(r#"
        fun main() {
            val mutable = mutableListOf(1, 2, 3)
            val asReadOnly = mutable as List<Int>
            println(asReadOnly.size)

            val backToMutable = mutable as? MutableList<Int>
            println(backToMutable != null)
            println(backToMutable?.size ?: -1)
        }
    "#);
    assert_eq!(out, &["3", "true", "3"]);
}

#[test]
fn test_readonly_list_safe_cast_to_mutable_is_nullable_failure() {
    let out = run_prints(r#"
        fun main() {
            val readonly: Any = listOf("x", "y")
            val casted = readonly as? MutableList<String>
            println(casted == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_is_check_on_generic_boxing_of_any() {
    let out = run_prints(r#"
        fun isStringList(value: Any): Boolean {
            return value is List<*>
        }

        fun main() {
            println(isStringList(listOf("a", "b", "c")))
            println(isStringList(10))
            val maybeList: Any? = null
            println(maybeList is List<*>)
        }
    "#);
    assert_eq!(out, &["true", "false", "false"]);
}

#[test]
fn test_type_projection_and_safe_cast_chain_in_expression() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = arrayOf(1, 2, 3)
            val values = (value as? IntArray) ?: intArrayOf(9, 8)
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["9,8"]);
}

#[test]
fn test_subject_when_type_dispatch_with_is() {
    let out = run_prints(r#"
        fun describe(value: Any): String {
            return when (value) {
                is Int -> if (value % 2 == 0) "even" else "odd"
                is String -> "len:" + value.length
                is Boolean -> "bool:" + (if (value) 1 else 0)
                else -> "other"
            }
        }

        fun main() {
            println(describe(4))
            println(describe("go"))
            println(describe(false))
            println(describe(3.14))
        }
    "#);
    assert_eq!(out, &["even", "len:2", "bool:0", "other"]);
}

#[test]
fn test_is_filtering_in_loop() {
    let out = run_prints(r#"
        fun main() {
            val items: Array<Any?> = arrayOf(1, "x", true, 2.5, null)
            var count = 0
            var stringLen = 0
            var boolSeen = false
            for (item in items) {
                if (item is Int) {
                    count += item
                } else if (item is String) {
                    stringLen = item.length
                } else if (item is Boolean) {
                    boolSeen = item
                }
            }
            println(count)
            println(stringLen)
            println(boolSeen)
        }
    "#);
    assert_eq!(out, &["1", "1", "true"]);
}

#[test]
fn test_smart_cast_remains_in_while_loop_body() {
    let out = run_prints(r#"
        fun main() {
            var value: Any? = "abc"
            while (value is String) {
                println(value.length)
                value = 0
            }
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_safe_cast_from_mixed_array_reference_fails() {
    let out = run_prints(r#"
        class Holder {
            val payload: Any = arrayOf(1, "x")
        }

        fun main() {
            val value: Any = Holder().payload
            val casted = value as? Array<Int>
            println(casted == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_unsafe_cast_to_wrong_class_is_caught() {
    let out = run_prints(r#"
        open class Animal
        class Dog : Animal()
        class Cat : Animal()

        fun main() {
            val value: Animal = Cat()
            try {
                val dog = value as Dog
                println(dog is Dog)
            } catch (e: Exception) {
                println("bad")
            }
        }
    "#);
    assert_eq!(out, &["bad"]);
}

#[test]
fn test_type_check_on_sealed_subtype() {
    let out = run_prints(r#"
        sealed class ResultState {
            class Ok(val value: Int) : ResultState()
            class Err(val message: String) : ResultState()
        }

        fun main() {
            val state: ResultState = ResultState.Ok(7)
            if (state is ResultState.Ok) {
                println(state.value)
            }
            val mapped = state as? ResultState.Err
            println(mapped == null)
        }
    "#);
    assert_eq!(out, &["7", "true"]);
}

#[test]
fn test_as_with_null_coalescing_is_non_throwing_for_nullable_source() {
    let out = run_prints(r#"
        fun extract(value: Any?): String {
            return (value as? String) ?: "missing"
        }

        fun main() {
            println(extract("kotlin"))
            println(extract(null))
        }
    "#);
    assert_eq!(out, &["kotlin", "missing"]);
}

#[test]
fn test_safe_cast_respects_function_type_arity_and_parameter_types() {
    let out = run_prints(r#"
        fun main() {
            val handler: Any = { value: Int -> value.toString() }
            val unary = handler as? (Int) -> String
            val binary = handler as? (Int, Int) -> String
            println(unary != null)
            println(binary == null)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_unsafe_cast_to_incompatible_function_type_is_caught() {
    let out = run_prints(r#"
        fun main() {
            val handler: Any = { value: String -> value + "!" }
            try {
                val bad = handler as (Int) -> String
                println("bad:" + bad(3))
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#);
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_is_check_for_number_interface_vs_concrete_numeric_type() {
    let out = run_prints(r#"
        fun main() {
            val intValue: Any = 12
            val longValue: Any = 12L
            val doubleValue: Any = 12.0

            println(intValue is Number)
            println(longValue is Int)
            println(doubleValue is Long)
            println(longValue is Number)
            println(intValue as? Int != null)
            println(longValue as? Int == null)
        }
    "#);
    assert_eq!(out, &["true", "false", "false", "true", "true", "true"]);
}

#[test]
fn test_as_list_to_mutable_list_preserves_reference_when_possible() {
    let out = run_prints(r#"
        fun main() {
            val readonly: List<Int> = listOf(1, 2, 3)
            try {
                val mutable = readonly as MutableList<Int>
                mutable.add(4)
                println("mutated")
            } catch (e: Exception) {
                println("rejected")
            }
        }
    "#);
    assert_eq!(out, &["rejected"]);
}

#[test]
fn test_casting_array_to_primitive_array_projection_is_type_checked() {
    let out = run_prints(r#"
        fun main() {
            val values: Any = arrayOf(5, 6, 7)
            val primitive = values as? IntArray
            val boxed = values as? Array<Int>
            println(primitive == null)
            println(boxed != null)
            println(boxed?.size)
        }
    "#);
    assert_eq!(out, &["true", "true", "3"]);
}

#[test]
fn test_smart_cast_lost_after_reassignment_in_the_same_scope() {
    let out = run_prints(r#"
        fun main() {
            var value: Any = "start"
            if (value is String) {
                println(value.length)
                value = 9
            }
            if (value is String) {
                println("after-string")
            } else {
                println("after-not-string")
            }
        }
    "#);
    assert_eq!(out, &["5", "after-not-string"]);
}

#[test]
fn test_casting_nullable_to_non_nullable_is_forced_and_throws() {
    let out = run_prints(r#"
        fun main() {
            val source: Any? = null
            val direct: String? = source as String?
            println(direct == null)

            try {
                val strict: String = source as String
                println(strict)
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#);
    assert_eq!(out, &["true", "caught"]);
}

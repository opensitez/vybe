use crate::helpers::run_prints;

#[test]
fn test_safe_call() {
    let out = run_prints(r#"
        class User(val name: String)

        fun main() {
            val u: User? = null
            println(u?.name ?: "No User")
        }
    "#);
    assert_eq!(out, &["No User"]);
}

#[test]
fn test_elvis_operator() {
    let out = run_prints(r#"
        fun main() {
            val name: String? = null
            val display = name ?: "default"
            println(display)
        }
    "#);
    assert_eq!(out, &["default"]);
}

#[test]
fn test_elvis_short_circuit_non_null() {
    let out = run_prints(r#"
        fun main() {
            val name: String? = "Kotlin"
            val display = name ?: "Fallback"
            println(display)
        }
    "#);
    assert_eq!(out, &["Kotlin"]);
}

#[test]
fn test_null_check_if_else() {
    let out = run_prints(r#"
        fun main() {
            val str: String? = null
            if (str != null) {
                println("valid")
            } else {
                println("null value")
            }
        }
    "#);
    assert_eq!(out, &["null value"]);
}

#[test]
fn test_non_null_variable() {
    let out = run_prints(r#"
        fun main() {
            val s: String = "Hello"
            println(s)
        }
    "#);
    assert_eq!(out, &["Hello"]);
}

#[test]
fn test_null_literal_assignment() {
    let out = run_prints(r#"
        fun main() {
            val s: String? = null
            println(s)
        }
    "#);
    assert_eq!(out, &["null"]);
}

#[test]
fn test_elvis_with_arithmetic() {
    let out = run_prints(r#"
        fun main() {
            val count: Int? = null
            val total = (count ?: 0) + 10
            println(total)
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_safe_call_on_present_object() {
    let out = run_prints(r#"
        class Item(val price: Int)

        fun main() {
            val item: Item? = Item(49)
            println(item?.price ?: 0)
        }
    "#);
    assert_eq!(out, &["49"]);
}

#[test]
fn test_null_assertion_operator() {
    let out = run_prints(r#"
        fun main() {
            val user: String? = "Vybe"
            val nonNull = user!!
            println(nonNull)
        }
    "#);
    assert_eq!(out, &["Vybe"]);
}

#[test]
fn test_null_assertion_failed() {
    let out = run_prints(r#"
        fun main() {
            val user: String? = null
            try {
                println(user!!)
            } catch (e: Exception) {
                println("null")
            }
        }
    "#);
    assert_eq!(out, &["null"]);
}

#[test]
fn test_safe_call_chained_property() {
    let out = run_prints(r#"
        class Profile {
            var name: String? = null
            fun label(): String {
                return name ?: "anon"
            }
        }

        fun main() {
            val p: Profile? = Profile()
            println(p?.label() ?: "none")
        }
    "#);
    assert_eq!(out, &["anon"]);
}

#[test]
fn test_nullable_with_elvis_and_binary_op() {
    let out = run_prints(r#"
        fun main() {
            val a: Int? = null
            val b: Int? = 4
            val left = a ?: 0
            val right = b ?: 0
            println(left + right)
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_safe_call_chain_with_method() {
    let out = run_prints(r#"
        class Node {
            fun name(): String {
                return "node"
            }
        }

        fun main() {
            val value: Node? = Node()
            println(value?.name() ?: "none")
        }
    "#);
    assert_eq!(out, &["node"]);
}

#[test]
fn test_nullable_with_while_loop() {
    let out = run_prints(r#"
        class Node {
            var next: Node? = null
        }

        fun main() {
            val head: Node? = Node()
            val result = if (head?.next == null) "end" else "mid"
            println(result)
        }
    "#);
    assert_eq!(out, &["end"]);
}

#[test]
fn test_nullable_chaining_with_safe_cast() {
    let out = run_prints(r#"
        open class Base
        class Child : Base()

        fun main() {
            val base: Base? = Child()
            val child = base as? Child
            if (child != null) {
                println("ok")
            }
        }
    "#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_safe_cast_failed_returns_null() {
    let out = run_prints(r#"
        open class Base
        class Child : Base()

        fun main() {
            val base: Base = Base()
            val child = base as? Child
            println(child == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_non_null_assertion_operator_success() {
    let out = run_prints(r#"
        fun main() {
            val text: String? = "value"
            val forced = text!!
            println(forced)
        }
    "#);
    assert_eq!(out, &["value"]);
}

#[test]
fn test_non_null_assertion_after_check() {
    let out = run_prints(r#"
        fun main() {
            val maybe: String? = "ok"
            if (maybe != null) {
                println(maybe!!)
            }
        }
    "#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_safe_call_with_null_subject() {
    let out = run_prints(r#"
        class Item {
            fun value(): Int = 7
        }

        fun main() {
            val item: Item? = null
            println(item?.value() ?: -1)
        }
    "#);
    assert_eq!(out, &["-1"]);
}

#[test]
fn test_null_elvis_on_boolean() {
    let out = run_prints(r#"
        fun main() {
            val flag: Boolean? = null
            val status = flag ?: false
            if (status) {
                println("on")
            } else {
                println("off")
            }
        }
    "#);
    assert_eq!(out, &["off"]);
}

#[test]
fn test_nullable_infix_safety() {
    let out = run_prints(r#"
        fun main() {
            val left: Int? = 3
            val right: Int? = null
            val leftValue = left ?: 0
            val rightValue = right ?: 10
            println(leftValue + rightValue)
        }
    "#);
    assert_eq!(out, &["13"]);
}

#[test]
fn test_nullable_array_of_objects() {
    let out = run_prints(r#"
        class Holder(val value: Int)

        fun main() {
            val first: Holder? = Holder(5)
            val second: Holder? = null
            println(first?.value ?: 0)
            println(second?.value ?: 9)
        }
    "#);
    assert_eq!(out, &["5", "9"]);
}

#[test]
fn test_nullability_or_return() {
    let out = run_prints(r#"
fun ensure(v: String?): String { return v ?: "empty" }; fun main() { println(ensure(null)); println(ensure("ok")) }
"#);
    assert_eq!(out, &["empty", "ok"]);
}

#[test]
fn test_nullability_safe_navigation_chain() {
    let out = run_prints(r#"
class Box(val value: Int); class Wrapper(val box: Box?); fun main() { val wrapped: Wrapper? = Wrapper(Box(4)); val fallback = wrapped?.box?.value ?: -1; println(fallback) }
"#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_nullability_nullable_parameter() {
    let out = run_prints(r#"
fun printIfNotNull(v: String?): Int { return if (v == null) 0 else v.length }; fun main() { println(printIfNotNull(null)); println(printIfNotNull("abc")) }
"#);
    assert_eq!(out, &["0", "3"]);
}

#[test]
fn test_nullability_not_null_assertion() {
    let out = run_prints(r#"
fun main() { val value: String? = "k"; println(value!!) }
"#);
    assert_eq!(out, &["k"]);
}

#[test]
fn test_nullability_caught_not_null_assertion() {
    let out = run_prints(r#"
fun main() { val value: String? = null; try { println(value!!) } catch (e: Exception) { println("fail") } }
"#);
    assert_eq!(out, &["fail"]);
}

#[test]
fn test_nullability_nonnull_in_function() {
    let out = run_prints(r#"
fun upper(value: String): String = value + value; fun main() { val s: String? = "abc"; println(upper(s!!)) }
"#);
    assert_eq!(out, &["abcabc"]);
}

#[test]
fn test_nullability_iff_check() {
    let out = run_prints(r#"
fun main() { val value: Int? = 10; if (value != null) { println(value) } else { println(0) } }
"#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_nullability_nested_optional() {
    let out = run_prints(r#"
fun main() { val a: String? = null; val b = a ?: ("z"); val c = b + "oo"; println(c) }
"#);
    assert_eq!(out, &["zoo"]);
}


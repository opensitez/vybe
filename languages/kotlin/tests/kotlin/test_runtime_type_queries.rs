use crate::helpers::run_prints;

#[test]
fn test_is_check_on_primitives_vs_wrappers() {
    let out = run_prints(r#"
        fun main() {
            val a: Any = 5
            val b: Any = "kotlin"
            println(a is Int)
            println(a is String)
            println(b is String)
            println(b is Number)
        }
    "#);
    assert_eq!(out, &["true", "false", "true", "false"]);
}

#[test]
fn test_safe_cast_with_as_question_mark() {
    let out = run_prints(r#"
        fun main() {
            val a: Any = "hello"
            val cast1 = a as? String
            val cast2 = a as? Int
            println(cast1 ?: "none")
            println(cast2?.toString() ?: "none")
        }
    "#);
    assert_eq!(out, &["hello", "none"]);
}

#[test]
fn test_unsafe_cast_fails_with_exception() {
    let out = run_prints(r#"
        fun main() {
            val a: Any = 7
            try {
                val s = a as String
                println(s)
            } catch (e: Exception) {
                println("err")
            }
        }
    "#);
    assert_eq!(out, &["err"]);
}

#[test]
fn test_when_type_guarding_with_is() {
    let out = run_prints(r#"
        fun main() {
            val values: List<Any> = listOf("k", 12, 3.4)
            val tags = values.map { value ->
                when (value) {
                    is Int -> "int"
                    is Double -> "double"
                    is String -> "string"
                    else -> "other"
                }
            }
            println(tags.joinToString(","))
        }
    "#);
    assert_eq!(out, &["string,int,double"]);
}

#[test]
fn test_nullable_is_checks() {
    let out = run_prints(r#"
        fun main() {
            val v: String? = null
            println(v is String)
            println(v == null)
            val w: Any? = null
            println(w is String?)
            println(w is Int?)
        }
    "#);
    assert_eq!(out, &["false", "true", "true", "true"]);
}

#[test]
fn test_smart_cast_after_is() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = "abc"
            if (value is String) {
                println(value.length)
            } else {
                println(0)
            }
            val boxed: Any = 123
            if (boxed is Int) {
                println(boxed + 1)
            }
        }
    "#);
    assert_eq!(out, &["3", "124"]);
}

#[test]
fn test_generic_type_checks_star_projection() {
    let out = run_prints(r#"
        fun main() {
            val values: Any = listOf(1, 2, 3)
            println(values is List<*>)
            println(values is Set<*>)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_sealed_type_exhaustive_when() {
    let out = run_prints(r#"
        sealed interface Node
        data class Leaf(val value: Int) : Node
        data class Branch(val left: Node, val right: Node) : Node

        fun classify(node: Node): String = when (node) {
            is Leaf -> "leaf"
            is Branch -> "branch"
        }

        fun main() {
            val a = Leaf(1)
            val b = Branch(Leaf(2), Leaf(3))
            println(classify(a))
            println(classify(b))
        }
    "#);
    assert_eq!(out, &["leaf", "branch"]);
}

#[test]
fn test_instanceof_interface_and_implementation() {
    let out = run_prints(r#"
        fun main() {
            interface Marker
            class A : Marker
            class B
            val a: Marker = A()
            val b: Any = B()
            println(a is Marker)
            println(b is Marker)
            println(a !is B)
            println(b is B)
        }
    "#);
    assert_eq!(out, &["true", "false", "true", "true"]);
}

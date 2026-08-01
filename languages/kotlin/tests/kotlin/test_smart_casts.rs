use crate::helpers::run_prints;

#[test]
fn test_is_operator_true_for_exact_type() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = "hello"
            println(value is String)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_is_operator_false_for_mismatched_type() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = 42
            println(value is String)
        }
    "#,
    );
    assert_eq!(out, &["false"]);
}

#[test]
fn test_not_is_operator_negation() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = 42
            println(value !is String)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_is_operator_with_interface_match() {
    let out = run_prints(
        r#"
        interface Pet
        class Dog : Pet
        class Car

        fun main() {
            val pet: Any = Dog()
            val other: Any = Car()
            println(pet is Pet)
            println(other is Pet)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_when_with_is_branches_uses_first_matching_branch() {
    let out = run_prints(
        r#"
        interface Shape
        class Circle : Shape
        class Square : Shape

        fun main() {
            val value: Shape = Circle()
            val label = when (value) {
                is Circle -> "circle"
                is Square -> "square"
                else -> "other"
            }
            println(label)
        }
    "#,
    );
    assert_eq!(out, &["circle"]);
}

#[test]
fn test_when_with_is_no_argument_evaluates_first_true() {
    let out = run_prints(
        r#"
        class Dog
        class Cat

        fun main() {
            val value: Any = Cat()
            val result = when {
                value is Dog -> "dog"
                value is Cat -> "cat"
                else -> "unknown"
            }
            println(result)
        }
    "#,
    );
    assert_eq!(out, &["cat"]);
}

#[test]
fn test_smart_cast_allows_string_specific_calls() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = "Rust"
            val upper = if (value is String) {
                value.uppercase()
            } else {
                "none"
            }
            println(upper)
        }
    "#,
    );
    assert_eq!(out, &["RUST"]);
}

#[test]
fn test_smart_cast_works_after_null_check() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any? = "x"
            if (value != null) {
                println(value is String)
                println(value.uppercase())
            } else {
                println(false)
            }
        }
    "#,
    );
    assert_eq!(out, &["true", "X"]);
}

#[test]
fn test_as_cast_success() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = "kotlin"
            val text = value as String
            println(text.length)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_as_cast_failure_throws_class_cast_exception() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                val value: Any = 7
                val text = value as String
                println(text)
            } catch (error: ClassCastException) {
                println("cast-failed")
            }
        }
    "#,
    );
    assert_eq!(out, &["cast-failed"]);
}

#[test]
fn test_safe_cast_returns_null_when_type_mismatch() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = 7
            val text: String? = value as? String
            println(text == null)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_safe_cast_success_preserves_value() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = "abc"
            val text: String? = value as? String
            println(text ?: "missing")
        }
    "#,
    );
    assert_eq!(out, &["abc"]);
}

#[test]
fn test_safe_cast_chain_on_nullable_source() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any? = null
            val text: String? = value as? String
            println(text ?: "empty")
        }
    "#,
    );
    assert_eq!(out, &["empty"]);
}

#[test]
fn test_type_test_with_boolean_and_and_guard() {
    let out = run_prints(
        r#"
        open class Base
        class Child : Base()

        fun main() {
            val value: Any = Child()
            println(value is Base && value is Child)
            println(!(value is Child && value is String))
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_type_test_on_inherited_class_chain() {
    let out = run_prints(
        r#"
        open class Node
        open class Container : Node()
        class Boxed : Container()

        fun main() {
            val value: Node = Boxed()
            println(value is Node)
            println(value is Container)
            println(value is Boxed)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_when_type_dispatch_with_three_branches() {
    let out = run_prints(
        r#"
        class A
        class B : A()
        class C : A()

        fun main() {
            val value: A = C()
            val label = when (value) {
                is B -> "b"
                is C -> "c"
                else -> "a"
            }
            println(label)
        }
    "#,
    );
    assert_eq!(out, &["c"]);
}

#[test]
fn test_looping_type_checks_in_collection() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf<Any>(1, "two", 3, "four")
            var strings = 0
            var totalLen = 0
            for (item in values) {
                if (item is String) {
                    strings += 1
                    totalLen += item.length
                }
            }
            println(strings)
            println(totalLen)
        }
    "#,
    );
    assert_eq!(out, &["2", "7"]);
}

#[test]
fn test_cast_then_cast_back_to_nullable() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = "hello"
            val direct: String? = value as String
            val again: String? = direct as? String
            println(direct == again)
            println(again?.length)
        }
    "#,
    );
    assert_eq!(out, &["true", "5"]);
}

#[test]
fn test_is_check_with_numeric_widening_not_applied() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = 7
            println(value is Int)
            println(value is Long)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_nested_smart_cast_after_outer_check() {
    let out = run_prints(
        r#"
        open class Base
        class Holder(val text: String) : Base()
        class Wrapper(val child: Base)

        fun main() {
            val value: Any = Wrapper(Holder("ok"))
            if (value is Wrapper && value.child is Holder) {
                println(value.child.text)
            } else {
                println("no")
            }
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_smart_cast_in_while_like_rewrite() {
    let out = run_prints(
        r#"
        fun toMessage(value: Any): String {
            var cursor: Any = value
            var result = ""
            if (cursor is String) {
                result = cursor + "!"
            }
            if (cursor is String) {
                result += " twice"
            }
            println(result)
            return result
        }

        fun main() {
            println(toMessage("x"))
            println(toMessage(4))
        }
    "#,
    );
    assert_eq!(out, &["x! twice", ""]);
}

#[test]
fn test_smart_cast_with_property_read_preserves_original_type_guard() {
    let out = run_prints(
        r#"
        class Holder {
            val value: String = "v"
        }

        fun main() {
            val value: Any = Holder()
            if (value is Holder) {
                println(value.value)
            }
            println(value is Holder)
        }
    "#,
    );
    assert_eq!(out, &["v", "true"]);
}

#[test]
fn test_is_check_false_on_null_rejected() {
    let out = run_prints(
        r#"
        fun main() {
            val value: String? = null
            println(value is String)
        }
    "#,
    );
    assert_eq!(out, &["false"]);
}

#[test]
fn test_not_is_then_else_split_path() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = 123
            val label = if (value !is String) {
                "not-string"
            } else {
                "is-string"
            }
            println(label)
        }
    "#,
    );
    assert_eq!(out, &["not-string"]);
}

#[test]
fn test_if_is_chain_different_type_branches() {
    let out = run_prints(
        r#"
        fun classify(value: Any): String {
            return if (value is Int) {
                "int-" + value
            } else if (value is String) {
                "string-" + value.length
            } else {
                "other"
            }
        }

        fun main() {
            println(classify(9))
            println(classify("abc"))
            println(classify(true))
        }
    "#,
    );
    assert_eq!(out, &["int-9", "string-3", "other"]);
}

#[test]
fn test_cast_to_common_supertype_then_refine_to_subtype() {
    let out = run_prints(
        r#"
        open class Base
        class Left : Base()
        class Right : Base()

        fun main() {
            val base: Base = Left()
            println(base is Left)
            println(base is Right)
            val value = base as? Left
            println(value != null)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_when_on_nullable_with_type_operators() {
    let out = run_prints(
        r#"
        fun main() {
            val values: List<Any?> = listOf(null, "a", 3)
            val labels = values.map { item ->
                when (item) {
                    null -> "null"
                    is String -> "str"
                    is Int -> "int"
                    else -> "other"
                }
            }
            println(labels.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["null,str,int"]);
}

#[test]
fn test_as_question_mark_on_incompatible_numeric_type_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = 9L
            val asInt: Int? = value as? Int
            val asLong: Long? = value as? Long
            println(asInt == null)
            println(asLong != null)
            println(asLong?.toString() ?: "none")
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "9"]);
}

use crate::helpers::run_prints;

#[test]
fn test_as_safe_cast_matches_type() {
    let out = run_prints(r#"
        fun main() {
            val value: Any? = "hello"
            val text = value as? String
            println(text)
        }
    "#);
    assert_eq!(out, &["hello"]);
}

#[test]
fn test_as_safe_cast_wrong_type_returns_null() {
    let out = run_prints(r#"
        fun main() {
            val value: Any? = 10
            val text = value as? String
            println(text == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_as_cast_wrong_type_throws() {
    let out = run_prints(r#"
        fun main() {
            val value: Any? = 10
            try {
                value as String
                println("ok")
            } catch (e: Exception) {
                println(e::class.simpleName)
            }
        }
    "#);
    assert_eq!(out, &["ClassCastException"]);
}

#[test]
fn test_is_check_true_and_false() {
    let out = run_prints(r#"
        fun main() {
            val x: Any = 7
            println(x is Int)
            println(x is String)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_is_not_check() {
    let out = run_prints(r#"
        fun main() {
            val x: Any = "kotlin"
            println(x !is Int)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_smart_cast_after_is_in_if() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = "hello"
            val out = if (value is String) {
                value.length
            } else {
                -1
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_smart_cast_with_when_subject() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = listOf(1, 2)
            val out = when (value) {
                is List<*> -> value.size
                else -> -1
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_list_generic_cast_with_star_projection() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = listOf("a", "b")
            val list = value as? List<*>
            println(list?.size ?: -1)
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_cast_to_nullable_type() {
    let out = run_prints(r#"
        fun main() {
            val value: String? = null
            val any: Any? = value
            val value2 = any as? String
            println(value2 == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_is_on_null_value() {
    let out = run_prints(r#"
        fun main() {
            val value: Any? = null
            println(value is String)
        }
    "#);
    assert_eq!(out, &["false"]);
}

#[test]
fn test_cast_primitive_to_boxed() {
    let out = run_prints(r#"
        fun main() {
            val value: Int = 3
            val boxed: Any = value
            val cast = boxed as Int
            println(cast + 1)
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_is_works_with_subtype() {
    let out = run_prints(r#"
        interface Animal { fun kind(): String }
        class Cat : Animal { override fun kind() = "cat" }
        class Dog : Animal { override fun kind() = "dog" }

        fun main() {
            val a: Animal = Cat()
            println(a is Animal)
            println(a is Cat)
            println(a is Dog)
            println(a.kind())
        }
    "#);
    assert_eq!(out, &["true", "true", "false", "cat"]);
}

#[test]
fn test_cast_by_interface_reference() {
    let out = run_prints(r#"
        interface I { fun value(): Int }
        class Impl(val value: Int) : I { override fun value(): Int = value }

        fun asI(a: Any?): Int {
            return (a as? I)?.value() ?: -1
        }

        fun main() {
            println(asI(Impl(7)))
            println(asI("x"))
        }
    "#);
    assert_eq!(out, &["7", "-1"]);
}

#[test]
fn test_map_cast_to_keyvalue_pair() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = mapOf("a" to 1)
            val cast = value as? Map<String, Int>
            println(cast?.size ?: -1)
        }
    "#);
    assert_eq!(out, &["1"]);
}

#[test]
fn test_array_cast_preserves_elements_type() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = arrayOf(1, 2, 3)
            val cast = value as? Array<Int>
            println(cast?.size ?: -1)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_array_as_list_safecast() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = intArrayOf(1, 2, 3)
            val cast = value as? List<Int>
            println(cast == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_function_type_is_check() {
    let out = run_prints(r#"
        fun op(v: Int): Int = v + 1
        fun main() {
            val value: Any = ::op
            println(value is (Int) -> Int)
            val f = value as? (Int) -> Int
            println(f?.invoke(3) ?: -1)
        }
    "#);
    assert_eq!(out, &["true", "4"]);
}

#[test]
fn test_lambda_is_check_in_conditional() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = { s: String -> s.uppercase() }
            if (value is (String) -> String) {
                println(value("ab"))
            } else {
                println("no")
            }
        }
    "#);
    assert_eq!(out, &["AB"]);
}

#[test]
fn test_nested_smart_cast_in_while() {
    let out = run_prints(r#"
        fun main() {
            var value: Any = "start"
            var out = 0
            while (value is String) {
                out = value.length
                value = 10
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_cast_nullable_union_like() {
    let out = run_prints(r#"
        fun main() {
            val a: Any? = null
            val b: String? = a as? String
            println(b == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_reified_smart_cast_not_used() {
    let out = run_prints(r#"
        inline fun <reified T> isType(value: Any): Boolean = value is T

        fun main() {
            println(isType<String>("x"))
            println(isType<Int>("x"))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_safe_cast_of_null_to_string() {
    let out = run_prints(r#"
        fun main() {
            val value: Any? = null
            val text = value as? String
            println(text)
        }
    "#);
    assert_eq!(out, &["null"]);
}

#[test]
fn test_require_cast_message() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = 1
            try {
                val text = value as String
                println(text)
            } catch (e: ClassCastException) {
                println("bad_cast")
            }
        }
    "#);
    assert_eq!(out, &["bad_cast"]);
}

#[test]
fn test_cast_chain_on_map_values() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = mapOf("a" to "v")
            val map = value as? Map<String, String>
            println(map?.get("a") ?: "none")
        }
    "#);
    assert_eq!(out, &["v"]);
}

#[test]
fn test_cast_chain_with_primitive_array() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = intArrayOf(1, 2, 3)
            val values = value as? IntArray
            println(values?.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_cast_to_superclass_interface() {
    let out = run_prints(r#"
        interface X
        class Y : X
        fun main() {
            val value: Any = Y()
            println(value is X)
            val cast = value as X
            println(cast is X)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_cast_with_nullable_generic_parameter() {
    let out = run_prints(r#"
        class Box<T>(val value: T)
        fun main() {
            val value: Any = Box("x")
            val cast = value as? Box<String>
            println(cast?.value ?: "none")
        }
    "#);
    assert_eq!(out, &["x"]);
}

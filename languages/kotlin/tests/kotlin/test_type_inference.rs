use crate::helpers::run_prints;

#[test]
fn test_type_inference_from_int_literal() {
    let out = run_prints(
        r#"
        fun main() {
            val n = 10
            println(n + 5)
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}

#[test]
fn test_type_inference_from_string_literal() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "abc"
            println(text.length)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_type_inference_from_constructor_call() {
    let out = run_prints(
        r#"
        class Box(val value: Int)
        fun main() {
            val b = Box(7)
            println(b.value)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_type_inference_in_list_literal() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 2, 3)
            println(values.sum())
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_type_inference_of_map_from_entries() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            println(map["b"])
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_type_inference_for_lambda_parameter() {
    let out = run_prints(
        r#"
        fun main() {
            val f = { x: Int -> x + 1 }
            println(f(3))
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_type_inference_for_lambda_without_annotations() {
    let out = run_prints(
        r#"
        fun main() {
            val add = { a: Int, b: Int -> a + b }
            println(add(1, 2))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_type_inference_with_target_typed_lambda() {
    let out = run_prints(
        r#"
        fun main() {
            val f: (String) -> Int = { it.length }
            println(f("ab"))
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_type_inference_in_destructuring() {
    let out = run_prints(
        r#"
        fun main() {
            val pair = Pair(1, "a")
            val (num, text) = pair
            println(num)
            println(text)
        }
    "#,
    );
    assert_eq!(out, &["1", "a"]);
}

#[test]
fn test_type_inference_with_nullability_guess() {
    let out = run_prints(
        r#"
        fun main() {
            val value: String? = "x"
            println(value?.length)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_type_inference_with_elvis() {
    let out = run_prints(
        r#"
        fun main() {
            val a: Int? = null
            val b: Int = a ?: 5
            println(b)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_type_inference_for_generic_function_output() {
    let out = run_prints(
        r#"
        fun <T> id(v: T): T = v
        fun main() {
            val x = id(4)
            val y = id("z")
            println(x)
            println(y)
        }
    "#,
    );
    assert_eq!(out, &["4", "z"]);
}

#[test]
fn test_type_inference_for_mutable_collection() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            values.add(4)
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["[1, 2, 3, 4]"]);
}

#[test]
fn test_type_inference_with_set() {
    let out = run_prints(
        r#"
        fun main() {
            val set = mutableSetOf(1, 2, 2)
            println(set.size)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_type_inference_in_when_branch_join() {
    let out = run_prints(
        r#"
        fun main() {
            val value = when (1) {
                1 -> "one"
                else -> "other"
            }
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["one"]);
}

#[test]
fn test_type_inference_with_if_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val value = if (true) 1 else 2
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_type_inference_for_for_loop_items() {
    let out = run_prints(
        r#"
        fun main() {
            var sum = 0
            for (v in listOf(1, 2, 3)) {
                sum += v
            }
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_type_inference_with_pair_list_map() {
    let out = run_prints(
        r#"
        fun main() {
            val pairs = listOf(Pair(1, "a"), Pair(2, "b"))
            val map = pairs.toMap()
            println(map[2])
        }
    "#,
    );
    assert_eq!(out, &["b"]);
}

#[test]
fn test_type_inference_in_higher_order_context() {
    let out = run_prints(
        r#"
        fun build(fn: (Int) -> Int): Int = fn(4)
        fun main() {
            val x = build { it + 1 }
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_type_inference_of_return_in_function() {
    let out = run_prints(
        r#"
        fun pick(flag: Boolean) = if (flag) 1 else 0
        fun main() {
            println(pick(true))
            println(pick(false))
        }
    "#,
    );
    assert_eq!(out, &["1", "0"]);
}

#[test]
fn test_type_inference_for_array_of_primitives() {
    let out = run_prints(
        r#"
        fun main() {
            val values = intArrayOf(1, 2, 3)
            println(values[1])
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_type_inference_boolean_logic() {
    let out = run_prints(
        r#"
        fun main() {
            val x = 1
            val y = x > 0
            println(y)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_type_inference_nested_generic_function() {
    let out = run_prints(
        r#"
        fun <T> box(v: T): List<T> = listOf(v)
        fun main() {
            val items = box(7)
            println(items[0])
            val text = box("a")
            println(text[0])
        }
    "#,
    );
    assert_eq!(out, &["7", "a"]);
}

#[test]
fn test_type_inference_of_caller_callee_with_any() {
    let out = run_prints(
        r#"
        fun asAny(v: Any) = v.toString()
        fun main() {
            val x = asAny(9)
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_type_inference_with_nullable_list() {
    let out = run_prints(
        r#"
        fun main() {
            val values: List<Int>? = listOf(1, 2)
            println(values?.size)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_type_inference_with_sequence_builder() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = generateSequence(1) { if (it < 4) it + 1 else null }
            println(seq.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_type_inference_in_try_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val v = try {
                val x = 1 / 1
                x
            } catch (e: Exception) {
                0
            }
            println(v)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_type_inference_local_return_in_lambda() {
    let out = run_prints(
        r#"
        fun main() {
            val f = fun(x: Int): Int {
                return x * 2
            }
            println(f(6))
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_type_inference_with_class_hierarchy() {
    let out = run_prints(
        r#"
        open class Base
        class Child : Base()
        fun id(base: Base): Base = base
        fun main() {
            val value = id(Child())
            println(value is Child)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_type_inference_map_filter_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val sum = listOf(1, 2, 3).filter { it > 1 }.sum()
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_type_inference_in_while_like_counter() {
    let out = run_prints(
        r#"
        fun main() {
            var i = 0
            var total = 0
            while (i < 3) {
                total += i
                i++
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_type_inference_for_char_range() {
    let out = run_prints(
        r#"
        fun main() {
            val letters = 'a'..'c'
            println(letters.count())
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_type_inference_with_data_class_copy() {
    let out = run_prints(
        r#"
        data class Box(val x: Int)
        fun main() {
            val b = Box(1)
            val c = b.copy(x = 2)
            println(c.x)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_type_inference_of_range_sum() {
    let out = run_prints(
        r#"
        fun main() {
            val values = 1..5
            println(values.sum())
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}

#[test]
fn test_type_inference_in_when_subject_cast() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = "abc"
            val len = when (val text = value) {
                is String -> text.length
                else -> 0
            }
            println(len)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_type_inference_in_function_call_chain() {
    let out = run_prints(
        r#"
        fun first(v: Int): Int = v + 1
        fun second(v: Int): Int = v * 2
        fun main() {
            val out = first(second(3))
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_type_inference_with_reified_like_simulation() {
    let out = run_prints(
        r#"
        inline fun <reified T> typeName(value: T): String = value!!::class.simpleName ?: ""
        fun main() {
            println(typeName(1))
            println(typeName("x"))
        }
    "#,
    );
    assert_eq!(out, &["Int", "String"]);
}

#[test]
fn test_type_inference_local_function_result() {
    let out = run_prints(
        r#"
        fun main() {
            fun f() = 42
            println(f())
        }
    "#,
    );
    assert_eq!(out, &["42"]);
}

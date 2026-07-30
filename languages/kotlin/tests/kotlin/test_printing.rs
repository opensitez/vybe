use crate::helpers::run_prints;

#[test]
fn test_printing_empty_string_results_in_no_visible_payload() {
    let out = run_prints(r#"
        fun main() {
            println("")
        }
    "#);
    assert_eq!(out, &[""]);
}

#[test]
fn test_printing_positive_integer() {
    let out = run_prints(r#"
        fun main() {
            println(12)
        }
    "#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_printing_negative_integer() {
    let out = run_prints(r#"
        fun main() {
            println(-7)
        }
    "#);
    assert_eq!(out, &["-7"]);
}

#[test]
fn test_printing_zero_and_sign_boundary() {
    let out = run_prints(r#"
        fun main() {
            println(0)
            println(+5)
            println(-0)
        }
    "#);
    assert_eq!(out, &["0", "5", "0"]);
}

#[test]
fn test_printing_boolean_values() {
    let out = run_prints(r#"
        fun main() {
            println(true)
            println(false)
            println(1 == 1)
            println(1 == 2)
        }
    "#);
    assert_eq!(out, &["true", "false", "true", "false"]);
}

#[test]
fn test_printing_double_precision_output() {
    let out = run_prints(r#"
        fun main() {
            println(3.5)
            println(0.125)
            println(-2.0)
        }
    "#);
    assert_eq!(out, &["3.5", "0.125", "-2"]);
}

#[test]
fn test_printing_special_double_values() {
    let out = run_prints(r#"
        fun main() {
            println((1.0 / 0.0).isInfinite())
            println((-1.0 / 0.0).isInfinite())
            println((0.0 / 0.0).isNaN())
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_printing_character_singleton() {
    let out = run_prints(r#"
        fun main() {
            println('A')
            println('9')
        }
    "#);
    assert_eq!(out, &["A", "9"]);
}

#[test]
fn test_printing_escape_sequence_newline_and_tab() {
    let out = run_prints(r#"
        fun main() {
            println("a\n b")
            println("x\t y")
        }
    "#);
    assert_eq!(out, &["a\n b", "x\t y"]);
}

#[test]
fn test_printing_escape_sequence_quotes_and_backslash() {
    let out = run_prints(r#"
        fun main() {
            println("\"q\"")
            println("a\\b")
            println("\u0041\u0062")
        }
    "#);
    assert_eq!(out, &["\"q\"", "a\\b", "Ab"]);
}

#[test]
fn test_printing_string_template_variable() {
    let out = run_prints(r#"
        fun main() {
            val name = "kotlin"
            println("language=$name")
        }
    "#);
    assert_eq!(out, &["language=kotlin"]);
}

#[test]
fn test_printing_string_template_expression() {
    let out = run_prints(r#"
        fun main() {
            val left = 2
            val right = 5
            println("sum=${left + right}")
            println("$left*$right=${left * right}")
        }
    "#);
    assert_eq!(out, &["sum=7", "2*5=10"]);
}

#[test]
fn test_printing_dollar_sign_escape() {
    let out = run_prints(r#"
        fun main() {
            println("cost: ${'$'}5")
        }
    "#);
    assert_eq!(out, &["cost: $5"]);
}

#[test]
fn test_printing_string_plus_concatenation() {
    let out = run_prints(r#"
        fun main() {
            val prefix = "a"
            println(prefix + " + " + 2)
        }
    "#);
    assert_eq!(out, &["a + 2"]);
}

#[test]
fn test_printing_raw_string_triple_quote_output() {
    let out = run_prints(r#"
        fun main() {
            val raw = """line1
line2
line3"""
            println(raw)
        }
    "#);
    assert_eq!(out, &["line1\nline2\nline3"]);
}

#[test]
fn test_printing_null_reference() {
    let out = run_prints(r#"
        fun main() {
            val value: String? = null
            println(value)
        }
    "#);
    assert_eq!(out, &["null"]);
}

#[test]
fn test_printing_array_to_string_form() {
    let out = run_prints(r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3)
            println(nums)
        }
    "#);
    assert_eq!(out, &["[1, 2, 3]"]);
}

#[test]
fn test_printing_list_to_string_form() {
    let out = run_prints(r#"
        fun main() {
            val nums = listOf(1, 2, 3)
            println(nums)
        }
    "#);
    assert_eq!(out, &["[1, 2, 3]"]);
}

#[test]
fn test_printing_set_to_string_form() {
    let out = run_prints(r#"
        fun main() {
            val nums = linkedSetOf(1, 2, 3)
            println(nums)
        }
    "#);
    assert_eq!(out, &["[1, 2, 3]"]);
}

#[test]
fn test_printing_map_to_string_form() {
    let out = run_prints(r#"
        fun main() {
            val data = linkedMapOf("a" to 1, "b" to 2)
            println(data)
        }
    "#);
    assert_eq!(out, &["{a=1, b=2}"]);
}

#[test]
fn test_printing_pair_to_string_form() {
    let out = run_prints(r#"
        fun main() {
            val point = Pair(3, "x")
            println(point)
        }
    "#);
    assert_eq!(out, &["(3, x)"]);
}

#[test]
fn test_printing_range_to_string_form() {
    let out = run_prints(r#"
        fun main() {
            println(1..3)
        }
    "#);
    assert_eq!(out, &["1..3"]);
}

#[test]
fn test_printing_range_for_loop_outputs_each_step() {
    let out = run_prints(r#"
        fun main() {
            var output = ""
            for (i in 1..3) {
                output += i.toString() + ","
            }
            println(output)
        }
    "#);
    assert_eq!(out, &["1,2,3,"]);
}

#[test]
fn test_printing_print_then_println_call_order() {
    let out = run_prints(r#"
        fun main() {
            print("first")
            print("-")
            println("second")
            print("tail")
        }
    "#);
    assert_eq!(out, &["first", "-", "second", "tail"]);
}

#[test]
fn test_printing_print_uses_argument_to_string_for_boolean_expression() {
    let out = run_prints(r#"
        fun main() {
            print(1 < 2)
            println(2 < 1)
            println(!(1 == 2))
        }
    "#);
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_printing_prints_custom_data_class_to_string() {
    let out = run_prints(r#"
        data class Box(val value: Int)

        fun main() {
            println(Box(42))
        }
    "#);
    assert_eq!(out, &["Box(value=42)"]);
}

#[test]
fn test_printing_prints_nested_data_class_hierarchy_to_string() {
    let out = run_prints(r#"
        data class Child(val id: Int)
        data class Wrapper(val child: Child)

        fun main() {
            println(Wrapper(Child(7)))
        }
    "#);
    assert_eq!(out, &["Wrapper(child=Child(id=7))"]);
}

#[test]
fn test_printing_nested_string_expression_and_number_format() {
    let out = run_prints(r#"
        fun main() {
            val size = 3
            val label = "count=$size"
            println(label)
            println("${label.length}")
        }
    "#);
    assert_eq!(out, &["count=3", "6"]);
}

#[test]
fn test_printing_boolean_composition_for_logging() {
    let out = run_prints(r#"
        fun main() {
            val active = true
            val ready = false
            println("active=${active && !ready}")
            println("ready=${!ready}")
            println("both=${active == true && ready == false}")
        }
    "#);
    assert_eq!(out, &["active=true", "ready=true", "both=true"]);
}

#[test]
fn test_printing_array_like_range_step_output() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(3, 6, 9)
            val output = values.joinToString("|")
            println(output)
        }
    "#);
    assert_eq!(out, &["3|6|9"]);
}

#[test]
fn test_printing_uses_nullability_projection_in_template() {
    let out = run_prints(r#"
        fun main() {
            val maybe: String? = null
            println("value=${maybe ?: \"missing\"}")
            val present: String? = "ok"
            println("value=${present ?: \"missing\"}")
        }
    "#);
    assert_eq!(out, &["value=missing", "value=ok"]);
}

#[test]
fn test_printing_joined_output_across_multiple_calls() {
    let out = run_prints(r#"
        fun main() {
            val a = listOf("a", "b", "c")
            for (value in a) {
                print(value)
                print("-")
            }
            println("done")
        }
    "#);
    assert_eq!(out, &["a", "-", "b", "-", "c", "-", "done"]);
}

#[test]
fn test_printing_printing_of_complex_math_expression() {
    let out = run_prints(r#"
        fun main() {
            val width = 3
            val height = 4
            println("area=${width * height}")
            println("perimeter=${2 * (width + height)}")
        }
    "#);
    assert_eq!(out, &["area=12", "perimeter=14"]);
}

#[test]
fn test_printing_printing_of_boolean_array_contents() {
    let out = run_prints(r#"
        fun main() {
            val flags = booleanArrayOf(true, false, true)
            println(flags.joinToString(","))
        }
    "#);
    assert_eq!(out, &["true,false,true"]);
}

#[test]
fn test_printing_printing_of_char_array_contents() {
    let out = run_prints(r#"
        fun main() {
            val chars = charArrayOf('a', 'b', 'c')
            println(chars.joinToString(""))
        }
    "#);
    assert_eq!(out, &["abc"]);
}

#[test]
fn test_printing_prints_multi_line_raw_template() {
    let out = run_prints(r#"
        fun main() {
            val label = "items"
            println("""$label:
  - one
  - two""")
        }
    "#);
    assert_eq!(out, &["items:\n  - one\n  - two"]);
}

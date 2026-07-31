use crate::helpers::run_prints;

#[test]
fn test_import_kotlin_math_abs() {
    let out = run_prints(r#"
        import kotlin.math.abs
        fun main() {
            println(abs(-5))
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_import_kotlin_math_sqrt() {
    let out = run_prints(r#"
        import kotlin.math.sqrt
        fun main() {
            println(sqrt(16.0).toInt())
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_import_alias_as() {
    let out = run_prints(r#"
        import kotlin.math.max as maxValue
        fun main() {
            println(maxValue(3, 7))
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_import_star_import_used() {
    let out = run_prints(r#"
        import kotlin.math.*
        import kotlin.math.PI
        fun main() {
            println(PI.toInt())
            println(round(3.4))
        }
    "#);
    assert_eq!(out, &["3", "3"]);
}

#[test]
fn test_import_qualified_call_after_rename() {
    let out = run_prints(r#"
        import java.lang.StringBuilder
        fun main() {
            val b = StringBuilder()
            b.append("a").append("b")
            println(b.toString())
        }
    "#);
    assert_eq!(out, &["ab"]);
}

#[test]
fn test_import_function_reference_from_imported_symbol() {
    let out = run_prints(r#"
        import kotlin.math.absoluteValue
        fun norm(v: Int) = v.absoluteValue
        fun main() {
            println(norm(-12))
        }
    "#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_import_local_alias_type() {
    let out = run_prints(r#"
        import kotlin.collections.HashMap as HM
        fun main() {
            val map = HM<String, Int>()
            map["x"] = 9
            println(map["x"])
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_import_multiple_symbols() {
    let out = run_prints(r#"
        import kotlin.math.abs
        import kotlin.math.roundToInt
        fun main() {
            println(abs(-2))
            println(2.8.roundToInt())
        }
    "#);
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_import_static_extension_member() {
    let out = run_prints(r#"
        import kotlin.text.capitalize
        fun main() {
            println("kotlin".capitalize())
        }
    "#);
    assert_eq!(out, &["Kotlin"]);
}

#[test]
fn test_import_nested_package_style() {
    let out = run_prints(r#"
        import kotlin.system.exitProcess
        fun status(v: Int): String = if (v > 0) "ok" else "bad"
        fun main() {
            println(status(3))
            println(status(0))
        }
    "#);
    assert_eq!(out, &["ok", "bad"]);
}

#[test]
fn test_import_java_util_arrays() {
    let out = run_prints(r#"
        import java.util.Arrays
        fun main() {
            val a = intArrayOf(3, 1, 2)
            Arrays.sort(a)
            println(a.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_import_shadowing_local_name() {
    let out = run_prints(r#"
        import kotlin.math.abs as absoluteValue
        fun main() {
            val abs = 99
            println(absoluteValue(-7))
            println(abs)
        }
    "#);
    assert_eq!(out, &["7", "99"]);
}

#[test]
fn test_imports_with_no_usage_still_parses() {
    let out = run_prints(r#"
        import kotlin.math.max
        import kotlin.math.min
        fun main() {
            println(max(9, 3))
            println(min(9, 3))
        }
    "#);
    assert_eq!(out, &["9", "3"]);
}

#[test]
fn test_import_rename_conflict_resolved() {
    let out = run_prints(r#"
        import kotlin.collections.setOf as setOfStrings
        fun main() {
            val s = setOfStrings("a", "b")
            println(s.size)
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_import_type_alias_name_collision_with_local() {
    let out = run_prints(r#"
        import kotlin.collections.List
        fun main() {
            val values: List<Int> = listOf(1, 2, 3)
            println(values.size)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_import_class_with_generic_param() {
    let out = run_prints(r#"
        import kotlin.collections.ArrayList
        fun main() {
            val list: ArrayList<Int> = ArrayList()
            list.add(1)
            list.add(2)
            println(list[0] + list[1])
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_import_function_and_method_combo() {
    let out = run_prints(r#"
        import kotlin.math.roundToInt
        fun main() {
            println(1.9.roundToInt())
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_import_array_functions() {
    let out = run_prints(r#"
        import kotlin.collections.maxOrNull
        fun main() {
            println(listOf(1, 5, 3).maxOrNull())
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_import_static_members_on_string() {
    let out = run_prints(r#"
        import kotlin.text.appendLine
        fun main() {
            val sb = StringBuilder()
            sb.appendLine("a")
            sb.appendLine("b")
            println(sb.toString().trim())
        }
    "#);
    assert_eq!(out, &["a
b"]);
}

#[test]
fn test_import_invalid_alias_not_allowed_is_not_compiled() {
    let out = run_prints(r#"
        import kotlin.math.abs as absolute
        fun main() {
            println(absolute(-4))
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_import_java_math_namespace() {
    let out = run_prints(r#"
        import java.lang.Math
        fun main() {
            println(Math.max(2, 8))
        }
    "#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_import_extension_scope_conflict() {
    let out = run_prints(r#"
        import kotlin.text.uppercase
        import kotlin.text.lowercase
        fun main() {
            val text = "Ab"
            println(text.uppercase())
            println(text.lowercase())
        }
    "#);
    assert_eq!(out, &["AB", "ab"]);
}

#[test]
fn test_import_package_objects_in_expression() {
    let out = run_prints(r#"
        import kotlin.math.sqrt
        import kotlin.math.PI
        fun main() {
            println((sqrt(PI) * 2).toInt())
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_import_after_class_definition() {
    let out = run_prints(r#"
        class Holder
        import kotlin.math.absoluteValue
        fun main() {
            println((-9).absoluteValue)
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_import_function_reference_from_stdlib() {
    let out = run_prints(r#"
        import kotlin.math.abs
        fun main() {
            val f = ::abs
            println(f(-10))
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_import_multiple_aliases_same_base() {
    let out = run_prints(r#"
        import kotlin.math.sqrt as sq
        import kotlin.math.sqrt as squareRoot
        fun main() {
            println(sq(9.0).toInt())
            println(squareRoot(16.0).toInt())
        }
    "#);
    assert_eq!(out, &["3", "4"]);
}

#[test]
fn test_imports_in_generated_sequence() {
    let out = run_prints(r#"
        import kotlin.math.max
        fun main() {
            val values = generateSequence(0) { it + 1 }.take(4).toList()
            println(values.maxOrNull())
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_import_unused_but_parsed() {
    let out = run_prints(r#"
        import kotlin.collections.HashSet
        import kotlin.collections.HashMap
        fun main() {
            val a = HashSet<Int>()
            val b = HashMap<String, Int>()
            a.add(1)
            b["x"] = 2
            println(a.size)
            println(b["x"])
        }
    "#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_import_array_of_functions() {
    let out = run_prints(r#"
        import kotlin.math.abs
        fun main() {
            val ops: Array<(Int) -> Int> = arrayOf({ abs(it) }, { it * 2 })
            println(ops[0](-3))
            println(ops[1](4))
        }
    "#);
    assert_eq!(out, &["3", "8"]);
}

#[test]
fn test_import_non_ascii_path_alias() {
    let out = run_prints(r#"
        import kotlin.collections.mutableListOf as listOfAlias
        fun main() {
            val values = listOfAlias(1, 2, 3)
            println(values.size)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_import_with_class_scope_not_allowed() {
    let out = run_prints(r#"
        class Host {
            import kotlin.math.abs
            fun norm(v: Int): Int = abs(v)
        }
        fun main() {
            println(Host().norm(-7))
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_import_string_builder_from_kotlin() {
    let out = run_prints(r#"
        import kotlin.text.StringBuilder
        fun main() {
            val b = StringBuilder()
            b.append("x").append("y")
            println(b.toString())
        }
    "#);
    assert_eq!(out, &["xy"]);
}

#[test]
fn test_import_sequence_extension_call() {
    let out = run_prints(r#"
        import kotlin.collections.reduce
        fun main() {
            val sum = listOf(1, 2, 3).reduce { a, b -> a + b }
            println(sum)
        }
    "#);
    assert_eq!(out, &["6"]);
}

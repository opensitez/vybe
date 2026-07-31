use crate::helpers::run_prints;

#[test]
fn test_block_shadowing_keeps_outer_value_after_inner() {
    let out = run_prints(r#"
        fun main() {
            val value = "outer"
            val inside = run {
                val value = "inner"
                value
            }
            println(inside)
            println(value)
        }
    "#);
    assert_eq!(out, &["inner", "outer"]);
}

#[test]
fn test_loop_variable_does_not_escape_to_outer() {
    let out = run_prints(r#"
        fun main() {
            var outer = 1
            for (outer in listOf(2, 3, 4)) {
                println(outer)
                break
            }
            println(outer)
        }
    "#);
    assert_eq!(out, &["2", "1"]);
}

#[test]
fn test_when_subject_shadowing() {
    let out = run_prints(r#"
        fun main() {
            val token = "root"
            val result = when (token) {
                "root" -> {
                    val token = 100
                    token
                }
                else -> token.length
            }
            println(result)
            println(token)
        }
    "#);
    assert_eq!(out, &["100", "root"]);
}

#[test]
fn test_function_parameter_shadows_outer_property() {
    let out = run_prints(r#"
        val label = "outer"
        fun labelValue(label: String): String {
            return label
        }

        fun main() {
            println(labelValue("inner"))
            println(label)
        }
    "#);
    assert_eq!(out, &["inner", "outer"]);
}

#[test]
fn test_lambda_parameter_shadows_outer_var() {
    let out = run_prints(r#"
        fun main() {
            val value = "outer"
            val transform = { value: Int -> "lambda:$value" }
            println(transform(7))
            println(value)
        }
    "#);
    assert_eq!(out, &["lambda:7", "outer"]);
}

#[test]
fn test_nested_lambda_shadowing_chain() {
    let out = run_prints(r#"
        fun main() {
            val prefix = "A"
            val f = { prefix: String ->
                { prefix: Int -> "${'$'}{prefix}_${'$'}{prefix + 1}" }
            }
            val g = f("B")
            println(g(3))
            println(prefix)
        }
    "#);
    assert_eq!(out, &["B_4", "A"]);
}

#[test]
fn test_catch_block_shadowing_catch_name() {
    let out = run_prints(r#"
        fun main() {
            val e = "outer"
            try {
                throw IllegalStateException("boom")
            } catch (e: Exception) {
                println(e.message)
            }
            println(e)
        }
    "#);
    assert_eq!(out, &["boom", "outer"]);
}

#[test]
fn test_property_shadowing_in_nested_class() {
    let out = run_prints(r#"
        open class Base(val value: String)
        class Holder(overrideValue: String) : Base("base") {
            val value = overrideValue
            fun show(): String {
                return value
            }
        }

        fun main() {
            val holder = Holder("inner")
            println(holder.show())
            println((holder as Base).value)
        }
    "#);
    assert_eq!(out, &["inner", "base"]);
}

#[test]
fn test_shadowing_inside_if_branches() {
    let out = run_prints(r#"
        fun main() {
            val value = 1
            val out = if (value == 1) {
                val value = "one"
                value
            } else {
                "other"
            }
            println(out)
            println(value)
        }
    "#);
    assert_eq!(out, &["one", "1"]);
}

#[test]
fn test_nested_if_shadowing_isolated() {
    let out = run_prints(r#"
        fun main() {
            val n = 10
            fun test(x: Int): Int {
                val n = x + 1
                return if (x > 5) {
                    val n = n + 5
                    n
                } else {
                    n
                }
            }
            println(test(6))
            println(n)
        }
    "#);
    assert_eq!(out, &["12", "10"]);
}

#[test]
fn test_shadowing_after_mutable_update() {
    let out = run_prints(r#"
        fun main() {
            var value = 1
            run {
                var value = value + 1
                println(value)
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["2", "1"]);
}

#[test]
fn test_method_parameter_and_member_shadowing() {
    let out = run_prints(r#"
        class Holder {
            val value: String = "member"
            fun label(value: String): String {
                return value
            }
        }

        fun main() {
            val holder = Holder()
            println(holder.label("param"))
            println(holder.value)
        }
    "#);
    assert_eq!(out, &["param", "member"]);
}

#[test]
fn test_shadowing_preserves_this_reference() {
    let out = run_prints(r#"
        class Box {
            val value = "box"
            fun run(): String {
                val value = "inner"
                return this.value
            }
        }
        fun main() {
            println(Box().run())
        }
    "#);
    assert_eq!(out, &["box"]);
}

#[test]
fn test_receiver_shadowing_in_with() {
    let out = run_prints(r#"
        data class Node(val label: String)
        fun main() {
            val node = Node("outer")
            val label = "local"
            val out = with(node) {
                val label = "with"
                label + "|" + this.label
            }
            println(out)
            println(label)
        }
    "#);
    assert_eq!(out, &["with|outer", "local"]);
}

#[test]
fn test_object_expression_shadows_var() {
    let out = run_prints(r#"
        fun main() {
            var tag = "outer"
            val obj = object {
                val tag = "inner"
                fun value(): String = tag
            }
            println(obj.value())
            println(tag)
        }
    "#);
    assert_eq!(out, &["inner", "outer"]);
}

#[test]
fn test_let_shadowing_chain() {
    let out = run_prints(r#"
        fun main() {
            var value = "outer"
            val result = value.let { value ->
                val value = value + ":inner"
                value
            }
            println(result)
            println(value)
        }
    "#);
    assert_eq!(out, &["outer:inner", "outer"]);
}

#[test]
fn test_for_each_lambda_parameter_shadowing_outer() {
    let out = run_prints(r#"
        fun main() {
            val value = "outer"
            val values = listOf("a", "b")
            values.forEach { value ->
                println(value)
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["a", "b", "outer"]);
}

#[test]
fn test_nested_local_function_shadowing() {
    let out = run_prints(r#"
        fun main() {
            fun outer(): String {
                val value = "outer"
                fun inner(): String {
                    val value = "inner"
                    return value
                }
                return "${'$'}{inner()}|${'$'}value"
            }
            println(outer())
        }
    "#);
    assert_eq!(out, &["inner|outer"]);
}

#[test]
fn test_shadowed_variable_in_map_lambda() {
    let out = run_prints(r#"
        fun main() {
            val value = 1
            val result = listOf(1, 2, 3).map { value -> value * 2 }
            println(result.joinToString(","))
            println(value)
        }
    "#);
    assert_eq!(out, &["2,4,6", "1"]);
}

#[test]
fn test_shadowing_in_destructuring() {
    let out = run_prints(r#"
        fun main() {
            val outer = "outer"
            val pair = Pair("inner", 1)
            val (value, count) = pair
            println(value)
            println(count)
            println(outer)
        }
    "#);
    assert_eq!(out, &["inner", "1", "outer"]);
}

#[test]
fn test_destructure_local_overwrites() {
    let out = run_prints(r#"
        fun main() {
            val value = "outer"
            val (value, count) = listOf("x", "y").withIndex().first()
            println(value)
            println(count)
            println("outer" )
        }
    "#);
    assert_eq!(out, &["x", "0", "outer"]);
}

#[test]
fn test_lambda_with_receiver_shadowing_receiver_property() {
    let out = run_prints(r#"
        class Profile(val name: String)
        fun main() {
            val name = "outer"
            val profile = Profile("inner")
            val out = profile.run {
                val name = this.name
                val innerName = "inner2"
                innerName
            }
            println(out)
            println(name)
        }
    "#);
    assert_eq!(out, &["inner2", "outer"]);
}

#[test]
fn test_class_property_shadowing_in_inheritance() {
    let out = run_prints(r#"
        open class Parent {
            val value = "parent"
        }
        class Child : Parent() {
            val value = "child"
            fun reveal() = super.value + ":" + value
        }
        fun main() {
            val c = Child()
            println(c.value)
            println(c.reveal())
        }
    "#);
    assert_eq!(out, &["child", "parent:child"]);
}

#[test]
fn test_shadowing_with_return_inside_run() {
    let out = run_prints(r#"
        fun main() {
            val marker = "outer"
            val result = run {
                val marker = "inner"
                marker
            }
            println(result)
            println(marker)
        }
    "#);
    assert_eq!(out, &["inner", "outer"]);
}

#[test]
fn test_shadowing_for_same_file_scope_not_persisting() {
    let out = run_prints(r#"
        fun main() {
            val marker = "x"
            fun first() { val marker = "y"; println(marker) }
            fun second() { println(marker) }
            first()
            second()
        }
    "#);
    assert_eq!(out, &["y", "x"]);
}

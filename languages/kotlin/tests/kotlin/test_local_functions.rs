use crate::helpers::run_prints;

#[test]
fn test_local_function_adds_values() {
    let out = run_prints(r#"
        fun main() {
            fun add(a: Int, b: Int): Int = a + b
            println(add(2, 3))
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_local_function_uses_outer_scope() {
    let out = run_prints(r#"
        fun main() {
            val base = 10
            fun scale(x: Int): Int = x * base
            println(scale(4))
        }
    "#);
    assert_eq!(out, &["40"]);
}

#[test]
fn test_nested_local_function_call_chain() {
    let out = run_prints(r#"
        fun main() {
            fun outer(x: Int): Int {
                fun inner(y: Int): Int = y + 1
                return inner(x) * 2
            }
            println(outer(7))
        }
    "#);
    assert_eq!(out, &["16"]);
}

#[test]
fn test_local_function_with_defaults() {
    let out = run_prints(r#"
        fun main() {
            fun greet(name: String, suffix: String = "!"): String = "hi " + name + suffix
            println(greet("kotlin"))
            println(greet("kotlin", "!!"))
        }
    "#);
    assert_eq!(out, &["hi kotlin!", "hi kotlin!!"]);
}

#[test]
fn test_local_function_inside_if_branch() {
    let out = run_prints(r#"
        fun main() {
            val isAdmin = true
            val value = if (isAdmin) {
                fun role(): String = "admin"
                role()
            } else {
                fun role(): String = "user"
                role()
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["admin"]);
}

#[test]
fn test_local_function_recursive_factorial() {
    let out = run_prints(r#"
        fun main() {
            fun fact(n: Int): Int {
                return if (n <= 1) 1 else n * fact(n - 1)
            }
            println(fact(5))
        }
    "#);
    assert_eq!(out, &["120"]);
}

#[test]
fn test_local_function_with_tailcall_style() {
    let out = run_prints(r#"
        fun main() {
            fun sum(n: Int, acc: Int = 0): Int {
                return if (n == 0) acc else sum(n - 1, acc + n)
            }
            println(sum(4))
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_local_function_with_varargs_input() {
    let out = run_prints(r#"
        fun main() {
            fun join(prefix: String, vararg parts: Int): String {
                var out = prefix
                for (part in parts) {
                    out += ":" + part.toString()
                }
                return out
            }
            println(join("a", 1, 2, 3))
        }
    "#);
    assert_eq!(out, &["a:1:2:3"]);
}

#[test]
fn test_local_function_returning_lambda() {
    let out = run_prints(r#"
        fun main() {
            fun make(prefix: String): (Int) -> String {
                return { value -> "$prefix$value" }
            }
            val f = make("x")
            println(f(9))
        }
    "#);
    assert_eq!(out, &["x9"]);
}

#[test]
fn test_local_function_accepts_function_argument() {
    let out = run_prints(r#"
        fun main() {
            fun applyAndDescribe(value: Int, transform: (Int) -> Int): Int {
                return transform(value)
            }
            fun map(v: Int): Int = v + 1
            println(applyAndDescribe(4, ::map))
            println(applyAndDescribe(4) { it * 2 })
        }
    "#);
    assert_eq!(out, &["5", "8"]);
}

#[test]
fn test_local_function_called_from_local_function() {
    let out = run_prints(r#"
        fun main() {
            fun outer(x: Int): Int {
                fun plusOne(v: Int): Int = v + 1
                fun plusTwo(v: Int): Int = plusOne(v) + 1
                return plusTwo(x)
            }
            println(outer(3))
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_local_function_with_multiple_return_paths() {
    let out = run_prints(r#"
        fun main() {
            fun classify(v: Int): String {
                if (v < 0) return "neg"
                if (v == 0) return "zero"
                return "pos"
            }
            println(classify(-1))
            println(classify(0))
            println(classify(1))
        }
    "#);
    assert_eq!(out, &["neg", "zero", "pos"]);
}

#[test]
fn test_local_function_in_while_like_rewrite() {
    let out = run_prints(r#"
        fun main() {
            fun next(base: Int): Int = base + 1
            var i = 0
            var sum = 0
            while (i < 4) {
                sum += next(i)
                i++
            }
            println(sum)
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_local_function_name_hides_top_level_function() {
    let out = run_prints(r#"
        fun format(value: Int): Int = value

        fun main() {
            fun format(value: String): String = value + "!"
            println(format("x"))
            println(format(3))
        }
    "#);
    assert_eq!(out, &["x!", "3"]);
}

#[test]
fn test_local_function_reassignment_not_allowed_and_compiler_checks() {
    let out = run_prints(r#"
        fun main() {
            fun square(x: Int): Int = x * x
            try {
                println(square(4))
            } catch (error: Exception) {
                println("bad")
            }
        }
    "#);
    assert_eq!(out, &["16"]);
}

#[test]
fn test_local_function_uses_named_arguments() {
    let out = run_prints(r#"
        fun main() {
            fun build(a: Int, b: Int = 2, c: Int = 3): Int = a + b + c
            println(build(a = 1))
            println(build(a = 1, c = 10))
        }
    "#);
    assert_eq!(out, &["6", "13"]);
}

#[test]
fn test_local_function_in_for_loop_body() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(1, 2, 3)
            var total = 0
            fun add(v: Int) {
                total += v
            }
            for (value in values) {
                add(value)
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_local_function_with_mutable_capture() {
    let out = run_prints(r#"
        fun main() {
            var count = 0
            fun inc(amount: Int) {
                count += amount
            }
            inc(2)
            inc(3)
            println(count)
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_local_function_return_type_inference() {
    let out = run_prints(r#"
        fun main() {
            fun add(a: Int, b: Int) = a + b
            val value = add(3, 4)
            println(value)
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_local_function_inside_class_method() {
    let out = run_prints(r#"
        class Engine {
            fun execute(base: Int): Int {
                fun bump(v: Int): Int = v + 1
                return bump(base)
            }
        }

        fun main() {
            println(Engine().execute(7))
        }
    "#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_local_function_shadowing_outer_variable_name() {
    let out = run_prints(r#"
        fun main() {
            val value = 2
            fun compute(value: Int): Int = value + 1
            println(compute(10))
            println(value)
        }
    "#);
    assert_eq!(out, &["11", "2"]);
}

#[test]
fn test_local_function_in_try_catch_blocks() {
    let out = run_prints(r#"
        fun main() {
            try {
                fun parse(v: String): Int {
                    if (v.length == 0) throw RuntimeException("empty")
                    return v.toInt()
                }
                println(parse("12"))
            } catch (error: RuntimeException) {
                println("bad")
            }
        }
    "#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_local_function_without_parameters() {
    let out = run_prints(r#"
        fun main() {
            var value = 0
            fun tick() { value += 1 }
            tick()
            tick()
            println(value)
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_local_function_with_local_classes() {
    let out = run_prints(r#"
        fun main() {
            fun make(): String {
                class Holder(val v: Int)
                return Holder(9).v.toString()
            }
            println(make())
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_local_function_nested_across_scopes() {
    let out = run_prints(r#"
        fun main() {
            fun outer(v: Int): Int {
                fun inner1(x: Int): Int = x + 1
                if (v > 0) {
                    fun inner2(y: Int): Int = inner1(y * 2)
                    return inner2(v)
                }
                return inner1(v)
            }
            println(outer(3))
            println(outer(0))
        }
    "#);
    assert_eq!(out, &["7", "1"]);
}

#[test]
fn test_local_function_compiles_with_boolean_logic() {
    let out = run_prints(r#"
        fun main() {
            fun isEven(v: Int): Boolean = (v % 2 == 0)
            fun describe(v: Int): String = if (isEven(v)) "even" else "odd"
            println(describe(4))
            println(describe(5))
        }
    "#);
    assert_eq!(out, &["even", "odd"]);
}

#[test]
fn test_local_function_uses_its_parameter_names_as_types() {
    let out = run_prints(r#"
        fun main() {
            fun compose(prefix: String, suffix: String): String {
                fun join(value: String): String = prefix + value + suffix
                return join("mid")
            }
            println(compose("<", ">"))
        }
    "#);
    assert_eq!(out, &["<mid>"]);
}

#[test]
fn test_local_function_isolation_between_calls() {
    let out = run_prints(r#"
        fun main() {
            fun runOnce(v: Int): Int {
                fun bump(x: Int): Int = x + 1
                return bump(v)
            }
            println(runOnce(1))
            println(runOnce(2))
        }
    "#);
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_local_function_with_overloaded_names() {
    let out = run_prints(r#"
        fun main() {
            fun parseText(value: String): String = "s:" + value
            fun parseText(value: Int): String = "i:" + value
            println(parseText("x"))
            println(parseText(8))
        }
    "#);
    assert_eq!(out, &["s:x", "i:8"]);
}

#[test]
fn test_local_function_in_try_finally_path() {
    let out = run_prints(r#"
        fun main() {
            var executed = false
            fun body(v: Int): Int {
                executed = true
                return v
            }
            val value = try {
                body(9)
            } finally {
                println(if (executed) "ok" else "missing")
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["ok", "9"]);
}

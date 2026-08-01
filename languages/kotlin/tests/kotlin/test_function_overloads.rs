use crate::helpers::run_prints;

#[test]
fn test_overload_by_argument_count() {
    let out = run_prints(
        r#"
        fun token(x: Int): String = "int:" + x
        fun token(x: String, y: String): String = x + y
        fun main() {
            println(token(4))
            println(token("a", "b"))
        }
    "#,
    );
    assert_eq!(out, &["int:4", "ab"]);
}

#[test]
fn test_overload_by_type() {
    let out = run_prints(
        r#"
        fun cast(v: Int): String = "I"
        fun cast(v: String): String = "S"
        fun main() {
            println(cast(3))
            println(cast("x"))
        }
    "#,
    );
    assert_eq!(out, &["I", "S"]);
}

#[test]
fn test_overload_with_default_and_no_default_call() {
    let out = run_prints(
        r#"
        fun mark(v: Int): String = "single"
        fun mark(v: Int, suffix: String = ""): String = "double:" + v + suffix
        fun main() {
            println(mark(1))
            println(mark(2, "ok"))
        }
    "#,
    );
    assert_eq!(out, &["double:1", "double:2ok"]);
}

#[test]
fn test_overload_boolean_and_int_prefers_exact() {
    let out = run_prints(
        r#"
        fun value(v: Int): String = "num"
        fun value(v: Boolean): String = "bool"
        fun main() {
            println(value(1))
            println(value(true))
        }
    "#,
    );
    assert_eq!(out, &["num", "bool"]);
}

#[test]
fn test_overload_array_vs_vararg() {
    let out = run_prints(
        r#"
        fun total(values: IntArray): Int = values.sum()
        fun total(a: Int, b: Int): Int = a + b
        fun main() {
            println(total(1, 2))
            println(total(intArrayOf(1, 2, 3)))
        }
    "#,
    );
    assert_eq!(out, &["3", "6"]);
}

#[test]
fn test_overload_with_nullable_non_nullable() {
    let out = run_prints(
        r#"
        fun show(v: String): String = "S:" + v
        fun show(v: String?): String = "N:" + (v ?: "nil")
        fun main() {
            println(show("x"))
        }
    "#,
    );
    assert_eq!(out, &["S:x"]);
}

#[test]
fn test_overload_on_list_and_set() {
    let out = run_prints(
        r#"
        fun size(items: List<Int>): Int = items.size
        fun size(items: Set<Int>): Int = items.size + 10
        fun main() {
            println(size(listOf(1, 2, 3)))
            println(size(setOf(1, 2, 3)))
        }
    "#,
    );
    assert_eq!(out, &["3", "13"]);
}

#[test]
fn test_overload_with_generics_single_signature() {
    let out = run_prints(
        r#"
        fun <T> join(values: List<T>): Int = values.size
        fun join(values: String): Int = values.length
        fun main() {
            println(join(listOf(1, 2)))
            println(join("ab"))
        }
    "#,
    );
    assert_eq!(out, &["2", "2"]);
}

#[test]
fn test_overload_rejects_ambiguous_not_tested_here() {
    let out = run_prints(
        r#"
        fun tag(v: Int, suffix: String = "a"): String = "i" + suffix
        fun tag(v: Double, suffix: String = "b"): String = "d" + suffix
        fun main() {
            println(tag(1))
            println(tag(1.0, "Z"))
        }
    "#,
    );
    assert_eq!(out, &["ia", "dZ"]);
}

#[test]
fn test_overload_member_method_dispatch() {
    let out = run_prints(
        r#"
        class Solver {
            fun eval(v: Int): Int = v + 1
            fun eval(v: String): String = v + "!"
        }
        fun main() {
            val s = Solver()
            println(s.eval(3))
            println(s.eval("x"))
        }
    "#,
    );
    assert_eq!(out, &["4", "x!"]);
}

#[test]
fn test_overload_top_level_and_local_same_name() {
    let out = run_prints(
        r#"
        fun ping(v: Int): String = "global"
        fun main() {
            fun ping(v: String): String = "local"
            println(ping(1))
            println(ping("a"))
        }
    "#,
    );
    assert_eq!(out, &["global", "local"]);
}

#[test]
fn test_overload_with_tailrec_and_overload() {
    let out = run_prints(
        r##"
        fun build(v: Int): Int = v
        fun build(v: Int, s: String): String = s + v
        fun main() {
            println(build(1))
            println(build(1, "#"))
        }
    "##,
    );
    assert_eq!(out, &["1", "#1"]);
}

#[test]
fn test_overload_boolean_and_unit() {
    let out = run_prints(
        r#"
        fun marker(v: Int): Int = v
        fun marker(v: Boolean): String = if (v) "on" else "off"
        fun main() {
            println(marker(2))
            println(marker(false))
        }
    "#,
    );
    assert_eq!(out, &["2", "off"]);
}

#[test]
fn test_overload_with_defaulted_tail_param() {
    let out = run_prints(
        r#"
        fun pair(a: Int): String = "single" + a
        fun pair(a: Int, b: Int = 1): String = "pair" + (a + b)
        fun main() {
            println(pair(3))
            println(pair(3, 2))
        }
    "#,
    );
    assert_eq!(out, &["single3", "pair5"]);
}

#[test]
fn test_overload_nested_call_resolution() {
    let out = run_prints(
        r#"
        fun call(v: Int): String = "i" + v
        fun call(v: Int, t: String): String = "it" + t
        fun main() {
            println(call(7))
            println(call(7, "x"))
        }
    "#,
    );
    assert_eq!(out, &["i7", "itx"]);
}

#[test]
fn test_overload_with_primitive_conversions_not_coercing() {
    let out = run_prints(
        r#"
        fun decode(v: Int): String = "i"
        fun decode(v: String): String = "s"
        fun main() {
            println(decode(1))
            println(decode("1"))
        }
    "#,
    );
    assert_eq!(out, &["i", "s"]);
}

#[test]
fn test_overload_in_generic_class_method() {
    let out = run_prints(
        r#"
        class Box {
            fun size(value: Int): String = "int"
            fun size(value: String): String = "str"
            fun size(value: List<Int>): String = "list"
        }
        fun main() {
            val b = Box()
            println(b.size(4))
            println(b.size("a"))
            println(b.size(listOf(1)))
        }
    "#,
    );
    assert_eq!(out, &["int", "str", "list"]);
}

#[test]
fn test_overload_with_vararg_and_single_arg() {
    let out = run_prints(
        r#"
        fun take(v: String): String = "one"
        fun take(vararg v: String): String = "many:" + v.size
        fun main() {
            println(take("a"))
            println(take("a", "b", "c"))
        }
    "#,
    );
    assert_eq!(out, &["one", "many:3"]);
}

#[test]
fn test_overload_same_name_across_extensions() {
    let out = run_prints(
        r#"
        class Host
        fun Host.label(): String = "host"
        fun Host.label(prefix: String): String = prefix + ":host"
        fun main() {
            val h = Host()
            println(h.label())
            println(h.label("x"))
        }
    "#,
    );
    assert_eq!(out, &["host", "x:host"]);
}

#[test]
fn test_overload_return_types_do_not_distinguish() {
    let out = run_prints(
        r#"
        fun value(v: Int): Int = v
        fun value(v: Int, b: Int): Int = v + b
        fun main() {
            println(value(2))
            println(value(2, 3))
        }
    "#,
    );
    assert_eq!(out, &["2", "5"]);
}

#[test]
fn test_overload_when_called_from_expression() {
    let out = run_prints(
        r#"
        fun calc(v: Int): Int = v * 2
        fun calc(v: String): String = v + "!"
        fun main() {
            println(calc(4) + 1)
            println(calc("x"))
        }
    "#,
    );
    assert_eq!(out, &["9", "x!"]);
}

#[test]
fn test_overload_with_receiver_chains() {
    let out = run_prints(
        r#"
        class Maker {
            fun build(v: Int): Int = v
            fun build(v: Int, tag: String): String = tag + v
        }
        fun main() {
            val out = Maker().run {
                build(1).toString() + "," + build(1, "x")
            }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["1,x1"]);
}

#[test]
fn test_overload_operator_style_names() {
    let out = run_prints(
        r#"
        fun plus(a: Int): Int = a
        fun plus(a: String): String = a
        fun main() {
            println(plus(5))
            println(plus("y"))
        }
    "#,
    );
    assert_eq!(out, &["5", "y"]);
}

#[test]
fn test_overload_with_lambda_parameters() {
    let out = run_prints(
        r#"
        fun compute(v: Int, f: (Int) -> Int): Int = f(v)
        fun compute(v: String, f: (String) -> String): String = f(v)
        fun main() {
            println(compute(3) { it + 1 })
            println(compute("x") { it + "!" })
        }
    "#,
    );
    assert_eq!(out, &["4", "x!"]);
}

#[test]
fn test_overload_with_defaulted_second_parameter() {
    let out = run_prints(
        r#"
        fun ping(v: Int): String = "solo"
        fun ping(v: Int, label: String = "ok"): String = v.toString() + label
        fun main() {
            println(ping(1))
            println(ping(1, "x"))
        }
    "#,
    );
    assert_eq!(out, &["solo", "1x"]);
}

#[test]
fn test_overload_longer_parameter_list() {
    let out = run_prints(
        r#"
        fun merge(a: Int): String = "a"
        fun merge(a: Int, b: Int): String = "ab"
        fun merge(a: Int, b: Int, c: Int): String = "abc"
        fun main() {
            println(merge(1))
            println(merge(1, 2))
            println(merge(1, 2, 3))
        }
    "#,
    );
    assert_eq!(out, &["a", "ab", "abc"]);
}

#[test]
fn test_overload_nested_call_with_same_name_and_return() {
    let out = run_prints(
        r#"
        fun wrap(v: Int): Int = v + 1
        fun wrap(v: Int, depth: Int): Int = v + depth
        fun main() {
            fun run(value: Int): Int = wrap(value)
            println(run(1))
            println(wrap(1, 9))
        }
    "#,
    );
    assert_eq!(out, &["2", "10"]);
}

#[test]
fn test_overload_with_pair_inputs() {
    let out = run_prints(
        r#"
        fun shape(v: Pair<Int, Int>): String = "pair"
        fun shape(v: Pair<String, String>): String = "sPair"
        fun main() {
            println(shape(Pair(1, 2)))
            println(shape(Pair("a", "b")))
        }
    "#,
    );
    assert_eq!(out, &["pair", "sPair"]);
}

#[test]
fn test_overload_with_nullable_and_nonnull() {
    let out = run_prints(
        r#"
        fun show(v: String): String = "NN"
        fun show(v: String?): String = "NULL"
        fun main() {
            println(show("x"))
            println(show(null))
        }
    "#,
    );
    assert_eq!(out, &["NN", "NULL"]);
}

#[test]
fn test_overload_boolean_vs_unit_not_possible() {
    let out = run_prints(
        r#"
        fun ping(v: Int): String = "i"
        fun ping(v: Boolean, force: Int = 0): String = "b" + force
        fun main() {
            println(ping(1))
            println(ping(true))
            println(ping(false, 2))
        }
    "#,
    );
    assert_eq!(out, &["i", "b0", "b2"]);
}

#[test]
fn test_overload_on_nested_type_shape() {
    let out = run_prints(
        r#"
        fun parse(v: Int): String = "int"
        fun parse(v: Any): String = "any"
        open class A
        class B : A()
        fun main() {
            println(parse(1))
            println(parse(B()))
        }
    "#,
    );
    assert_eq!(out, &["int", "any"]);
}

#[test]
fn test_overload_compound_expression_dispatch() {
    let out = run_prints(
        r##"
        fun format(v: Int, tag: String = "i"): String = tag + v
        fun format(v: String, tag: String = "s"): String = tag + v
        fun main() {
            println(format(4))
            println(format("x", "#"))
        }
    "##,
    );
    assert_eq!(out, &["i4", "#x"]);
}

#[test]
fn test_overload_with_inheritance_parameter_types() {
    let out = run_prints(
        r#"
        open class Node
        class Child : Node()
        class Other : Node()
        fun visit(v: Node): String = "node"
        fun visit(v: Child): String = "child"
        fun main() {
            println(visit(Child()))
            println(visit(Other()))
        }
    "#,
    );
    assert_eq!(out, &["child", "node"]);
}

#[test]
fn test_overload_in_ternary_like_selection() {
    let out = run_prints(
        r#"
        fun convert(v: Int): Int = v
        fun convert(v: String): Int = v.length
        fun pick(flag: Boolean, value: Int): Int = if (flag) convert(value) else convert(value.toString())
        fun main() {
            println(pick(true, 3))
            println(pick(false, 3))
        }
    "#,
    );
    assert_eq!(out, &["3", "1"]);
}

#[test]
fn test_overload_with_no_argument_and_named_defaults() {
    let out = run_prints(
        r#"
        fun flag(v: Int = 1): String = "n" + v
        fun flag(v: String = "s"): String = "s" + v
        fun main() {
            println(flag())
            println(flag(2))
            println(flag(v = "x"))
        }
    "#,
    );
    assert_eq!(out, &["n1", "n2", "sx"]);
}

#[test]
fn test_overload_on_member_reference() {
    let out = run_prints(
        r#"
        class Ops {
            fun resolve(v: Int): String = "i"
            fun resolve(v: String): String = "s"
            fun use(): String {
                val fInt = this::resolve
                return fInt(3) + "," + fInt("x")
            }
        }
        fun main() {
            println(Ops().use())
        }
    "#,
    );
    assert_eq!(out, &["i,s"]);
}

#[test]
fn test_overload_with_lambda_argument_ordering() {
    let out = run_prints(
        r#"
        fun exec(v: Int, f: () -> String): String = "i:" + f()
        fun exec(v: String, f: () -> String): String = "s:" + f()
        fun main() {
            println(exec(1) { "x" })
            println(exec("y") { "z" })
        }
    "#,
    );
    assert_eq!(out, &["i:x", "s:z"]);
}

#[test]
fn test_overload_for_class_and_top_level_names() {
    let out = run_prints(
        r#"
        fun echo(v: Int): String = "top" + v
        class Host {
            fun echo(v: Int): String = "member" + v
        }
        fun main() {
            val h = Host()
            println(echo(1))
            println(h.echo(2))
        }
    "#,
    );
    assert_eq!(out, &["top1", "member2"]);
}

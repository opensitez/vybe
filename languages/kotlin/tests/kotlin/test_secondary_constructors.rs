use crate::helpers::run_prints;

#[test]
fn test_secondary_constructor() {
    let out = run_prints(
        r#"
        class Person {
            val name: String
            constructor(name: String) {
                this.name = name
            }
        }

        fun main() {
            val p = Person("Alice")
            println(p.name)
        }
    "#,
    );
    assert_eq!(out, &["Alice"]);
}

#[test]
fn test_secondary_constructor_chain() {
    let out = run_prints(
        r#"
        class Rectangle {
            val width: Int
            val height: Int

            constructor(side: Int) : this(side, side) {
                println("square")
            }

            constructor(width: Int, height: Int) {
                this.width = width
                this.height = height
            }
        }

        fun main() {
            val square = Rectangle(3)
            println(square.width)
            println(square.height)
        }
    "#,
    );
    assert_eq!(out, &["square", "3", "3"]);
}

#[test]
fn test_constructor_property_assignment() {
    let out = run_prints(
        r#"
        class Rectangle {
            val width: Int
            val height: Int

            constructor(width: Int, height: Int) {
                this.width = width
                this.height = height
            }

            constructor(size: Int) : this(size, size) {
                println("square")
            }
        }

        fun main() {
            val r1 = Rectangle(2, 3)
            val r2 = Rectangle(4)
            println(r1.width)
            println(r1.height)
            println(r2.width)
            println(r2.height)
        }
    "#,
    );
    assert_eq!(out, &["square", "2", "3", "4", "4"]);
}

#[test]
fn test_constructor_chain_with_this() {
    let out = run_prints(
        r#"
        class Margin {
            val top: Int
            val right: Int
            val bottom: Int
            val left: Int

            constructor(all: Int) : this(all, all, all, all) {
                println("all")
            }

            constructor(top: Int, right: Int, bottom: Int, left: Int) {
                this.top = top
                this.right = right
                this.bottom = bottom
                this.left = left
            }
        }

        fun main() {
            val m = Margin(7)
            println(m.top)
            println(m.right)
            println(m.bottom)
            println(m.left)
        }
    "#,
    );
    assert_eq!(out, &["all", "7", "7", "7", "7"]);
}

#[test]
fn test_constructor_super_call() {
    let out = run_prints(
        r#"
        open class Animal(val name: String)

        class Dog : Animal {
            val age: Int

            constructor(name: String, age: Int) : super(name) {
                this.age = age
            }

            constructor(name: String) : this(name, 1)
        }

        fun main() {
            val a = Dog("Rex", 5)
            val b = Dog("Buddy")
            println(a.name)
            println(a.age)
            println(b.name)
            println(b.age)
        }
    "#,
    );
    assert_eq!(out, &["Rex", "5", "Buddy", "1"]);
}

#[test]
fn test_constructor_default_behavior() {
    let out = run_prints(
        r#"
        class Timer {
            val seconds: Int

            constructor() {
                this.seconds = 0
            }

            constructor(start: Int) {
                this.seconds = start
            }
        }

        fun main() {
            println(Timer().seconds)
            println(Timer(30).seconds)
        }
    "#,
    );
    assert_eq!(out, &["0", "30"]);
}

#[test]
fn test_constructor_three_levels() {
    let out = run_prints(
        r#"
        class Scale {
            val unit: Int

            constructor() {
                this.unit = 1
            }

            constructor(value: Int) : this() {
                println("scaled")
            }

            constructor(value: Int, factor: Int) : this(value) {
                println(value * factor)
            }
        }

        fun main() {
            Scale()
            Scale(4)
            Scale(5, 2)
        }
    "#,
    );
    assert_eq!(out, &["scaled", "scaled", "10"]);
}

#[test]
fn test_secondary_no_default_init() {
    let out = run_prints(
        r#"
        class A {
            val value: Int
            constructor(v: Int) {
                this.value = v
            }
        }

        fun main() {
            println(A(3).value)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_secondary_double_constructor() {
    let out = run_prints(
        r#"
        class B {
            val a: Int
            val b: Int

            constructor(a: Int) {
                this.a = a
                this.b = a
            }

            constructor(a: Int, b: Int) : this(a) {
                this.b = b
            }
        }

        fun main() {
            val x = B(2, 9)
            println(x.a)
            println(x.b)
        }
    "#,
    );
    assert_eq!(out, &["2", "9"]);
}

#[test]
fn test_secondary_three_level_chain() {
    let out = run_prints(
        r#"
        class C {
            val value: Int

            constructor() {
                this.value = 0
            }

            constructor(v: Int) : this() {
                this.value = v
            }

            constructor(v: Int, extra: Int) : this(v) {
                this.value = v + extra
            }
        }

        fun main() {
            println(C().value)
            println(C(4).value)
            println(C(4, 5).value)
        }
    "#,
    );
    assert_eq!(out, &["0", "4", "9"]);
}

#[test]
fn test_secondary_with_super_base() {
    let out = run_prints(
        r#"
        open class P(val id: Int)

        class C : P {
            val tag: Int

            constructor() : super(1) {
                this.tag = 2
            }

            constructor(multiplier: Int) : super(multiplier) {
                this.tag = multiplier * 2
            }
        }

        fun main() {
            println(C().id)
            println(C(3).id)
            println(C(3).tag)
        }
    "#,
    );
    assert_eq!(out, &["1", "3", "6"]);
}

#[test]
fn test_secondary_printing_chain() {
    let out = run_prints(
        r#"
        class Trace {
            val value: Int

            constructor() {
                this.value = 0
                println("zero")
            }

            constructor(v: Int) : this() {
                println(v)
                this.value = v
            }
        }

        fun main() {
            Trace()
            Trace(5)
        }
    "#,
    );
    assert_eq!(out, &["zero", "5"]);
}

#[test]
fn test_secondary_companion_use() {
    let out = run_prints(
        r#"
        class Factory {
            val size: Int
            companion object {
                fun create(v: Int): Int {
                    return v * 2
                }
            }

            constructor(v: Int) {
                this.size = v
            }
        }

        fun main() {
            println(Factory.create(4))
            println(Factory.create(Factory(6).size))
        }
    "#,
    );
    assert_eq!(out, &["8", "12"]);
}

#[test]
fn test_secondary_chain_with_defaults() {
    let out = run_prints(
        r#"
        class Packet {
            val x: Int

            constructor() {
                this.x = 0
            }

            constructor(v: Int) : this() {
                this.x = v
            }

            constructor(v: Int, d: Int, e: Int) : this(v) {
                this.x = v + d + e
            }
        }

        fun main() {
            val p = Packet(2, 3, 4)
            println(p.x)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_secondary_reassigning_value() {
    let out = run_prints(
        r#"
        class Counter {
            var value: Int

            constructor() {
                this.value = 1
            }

            constructor(value: Int, double: Boolean) : this(value) {
                if (double) {
                    this.value = value * 2
                } else {
                    this.value = value
                }
            }
        }

        fun main() {
            println(Counter(3, true).value)
            println(Counter(3, false).value)
        }
    "#,
    );
    assert_eq!(out, &["6", "3"]);
}

#[test]
fn test_secondary_empty_parameter() {
    let out = run_prints(
        r#"
        class Label {
            val value: String

            constructor(text: String) {
                this.value = text
            }

            constructor() : this("none")
        }

        fun main() {
            println(Label().value)
            println(Label("yes").value)
        }
    "#,
    );
    assert_eq!(out, &["none", "yes"]);
}

#[test]
fn test_secondary_with_mutable_property() {
    let out = run_prints(
        r#"
        class State {
            var total: Int

            constructor() {
                this.total = 0
            }

            constructor(init: Int) : this() {
                this.total = init
            }

            constructor(init: Int, add: Int) : this(init) {
                this.total = this.total + add
            }
        }

        fun main() {
            val s = State(4, 5)
            println(s.total)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_secondary_noisy_chain() {
    let out = run_prints(
        r#"
        class Log {
            val step: Int

            constructor() {
                this.step = 0
            }

            constructor(v: Int) : this() {
                this.step = v
                println("s")
            }

            constructor(v: Int, extra: Int) : this(v) {
                println("e")
                this.step = v + extra
            }
        }

        fun main() {
            Log(2)
            Log(3, 4)
        }
    "#,
    );
    assert_eq!(out, &["s", "e"]);
}

#[test]
fn test_secondary_with_interface() {
    let out = run_prints(
        r#"
        interface Marker

        class D : Marker {
            val value: Int

            constructor() {
                this.value = 1
            }

            constructor(v: Int) : this() {
                this.value = v
            }
        }

        fun main() {
            println(D().value)
            println(D(8).value)
        }
    "#,
    );
    assert_eq!(out, &["1", "8"]);
}

#[test]
fn test_secondary_boolean_constructor() {
    let out = run_prints(
        r#"
        class Flag {
            val value: Boolean

            constructor() {
                this.value = false
            }

            constructor(flag: Boolean) : this() {
                this.value = flag
            }
        }

        fun main() {
            println(Flag().value)
            println(Flag(true).value)
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_secondary_chain_of_three() {
    let out = run_prints(
        r#"
        class Ring {
            val value: Int

            constructor() {
                this.value = 1
            }

            constructor(a: Int) : this() {
                this.value = a
            }

            constructor(a: Int, b: Int) : this(a) {
                this.value = a + b
            }

            constructor(a: Int, b: Int, c: Int) : this(a, b) {
                this.value = this.value + c
            }
        }

        fun main() {
            println(Ring(2, 3).value)
            println(Ring(2, 3, 4).value)
        }
    "#,
    );
    assert_eq!(out, &["5", "9"]);
}

#[test]
fn test_secondary_text_constructs() {
    let out = run_prints(
        r#"
        class Text {
            val value: String

            constructor(prefix: String) {
                this.value = prefix
            }

            constructor(prefix: String, suffix: String) : this(prefix) {
                this.value = prefix + suffix
            }
        }

        fun main() {
            println(Text("a").value)
            println(Text("a", "b").value)
        }
    "#,
    );
    assert_eq!(out, &["a", "ab"]);
}

#[test]
fn test_secondary_nullable_chain() {
    let out = run_prints(
        r#"
        class Value {
            val value: Int?

            constructor(v: Int) {
                this.value = v
            }

            constructor(flag: Boolean) : this(if (flag) 1 else 0)
        }

        fun main() {
            println(Value(true).value)
            println(Value(false).value)
        }
    "#,
    );
    assert_eq!(out, &["1", "0"]);
}

#[test]
fn test_secondary_float_constructor() {
    let out = run_prints(
        r#"
        class Box {
            val value: Int

            constructor() {
                this.value = 0
            }

            constructor(v: Double) : this() {
                this.value = v.toInt()
            }
        }

        fun main() {
            println(Box(4.9).value)
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_secondary_with_array_input() {
    let out = run_prints(
        r#"
        class Bucket {
            val size: Int

            constructor(values: Array<Int>) {
                this.size = values.size
            }

            constructor(a: Int) : this(arrayOf(a)) {}
        }

        fun main() {
            println(Bucket(5).size)
            println(Bucket(arrayOf(1, 2, 3)).size)
        }
    "#,
    );
    assert_eq!(out, &["1", "3"]);
}

#[test]
fn test_secondary_nested_class_chain() {
    let out = run_prints(
        r#"
        class Holder {
            val x: Int
            class Child

            constructor() {
                this.x = 1
            }

            constructor(v: Int) : this() {
                this.x = v
            }

            constructor(v: Int, child: Child) : this(v) {
                child
            }
        }

        fun main() {
            val h = Holder(7, Holder.Child())
            println(h.x)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_secondary_counter_reset() {
    let out = run_prints(
        r#"
        class Seq {
            val n: Int

            constructor() {
                this.n = 0
            }

            constructor(v: Int) : this() {
                this.n = v
            }
        }

        fun main() {
            println(Seq().n)
            println(Seq(12).n)
        }
    "#,
    );
    assert_eq!(out, &["0", "12"]);
}

#[test]
fn test_secondary_computed_value() {
    let out = run_prints(
        r#"
        class Compute {
            val value: Int

            constructor(base: Int) {
                this.value = base * 2
            }

            constructor(base: Int, factor: Int) : this(base) {
                this.value = base + factor
            }
        }

        fun main() {
            val c = Compute(4, 5)
            println(c.value)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_secondary_with_init_printing() {
    let out = run_prints(
        r#"
        class Track {
            val value: Int
            init {
                println("init")
            }

            constructor() {
                this.value = 0
            }

            constructor(v: Int) : this() {
                this.value = v
            }
        }

        fun main() {
            Track()
            Track(3)
        }
    "#,
    );
    assert_eq!(out, &["init", "init"]);
}

#[test]
fn test_secondary_constructor_and_mutability() {
    let out = run_prints(
        r#"
        class Counter {
            var value: Int

            constructor() {
                this.value = 0
            }

            constructor(v: Int) : this() {
                this.value = v
            }

            constructor(v: Int, inc: Int, dec: Int) : this(v, inc) {
                this.value = this.value + inc - dec
            }
        }

        fun main() {
            println(Counter(5, 1, 1).value)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_secondary_constructor_chain_side_effect_count() {
    let out = run_prints(
        r#"
        class SequenceTracker {
            val value: Int

            constructor() {
                println("base")
                this.value = 0
            }

            constructor(start: Int) : this() {
                println("fromStart")
                this.value = start
            }

            constructor(start: Int, step: Int) : this(start) {
                println("fromStep")
                this.value = start + step
            }
        }

        fun main() {
            println(SequenceTracker().value)
            println(SequenceTracker(3).value)
            println(SequenceTracker(3, 4).value)
        }
    "#,
    );
    assert_eq!(out, &["base", "0", "fromStart", "3", "fromStep", "7"]);
}

#[test]
fn test_secondary_constructor_accessing_other_constructor_results() {
    let out = run_prints(
        r#"
        class Range {
            val min: Int
            val max: Int

            constructor(value: Int) {
                this.min = value
                this.max = value
            }

            constructor(from: Int, to: Int) : this(from) {
                this.max = to
            }

            fun width(): Int = max - min
        }

        fun main() {
            val a = Range(4)
            val b = Range(2, 8)
            println(a.width())
            println(b.width())
        }
    "#,
    );
    assert_eq!(out, &["0", "6"]);
}

#[test]
fn test_secondary_constructor_reassigns_var_property() {
    let out = run_prints(
        r#"
        class Box {
            var value: Int

            constructor() {
                this.value = 10
            }

            constructor(input: Int) : this() {
                this.value += input
            }
        }

        fun main() {
            println(Box().value)
            println(Box(5).value)
        }
    "#,
    );
    assert_eq!(out, &["10", "15"]);
}

#[test]
fn test_secondary_constructor_with_nested_class_argument() {
    let out = run_prints(
        r#"
        class Host {
            val name: String
            class Config

            constructor(value: String) {
                this.name = value
            }

            constructor(config: Config, value: String) : this(value) {
                val used = config
                println(used is Config)
            }
        }

        fun main() {
            println(Host("root").name)
            println(Host(Host.Config(), "inner").name)
        }
    "#,
    );
    assert_eq!(out, &["true", "root", "inner"]);
}

#[test]
fn test_secondary_constructor_with_default_boolean_flag() {
    let out = run_prints(
        r#"
        class Marker {
            val active: Boolean
            val label: String

            constructor(label: String) {
                this.label = label
                this.active = false
            }

            constructor(label: String, active: Boolean) : this(label) {
                if (active) this.label = label + "!"
            }
        }

        fun main() {
            val a = Marker("x")
            val b = Marker("x", true)
            println(a.active)
            println(a.label)
            println(b.label)
        }
    "#,
    );
    assert_eq!(out, &["false", "x", "x!"]);
}

#[test]
fn test_secondary_constructor_readonly_property_preserved() {
    let out = run_prints(
        r#"
        class Metric {
            val total: Int
            var tag: String

            constructor(base: Int) {
                this.total = base
                this.tag = "base"
            }

            constructor(base: Int, tag: String) : this(base) {
                this.tag = tag
            }
        }

        fun main() {
            val one = Metric(3)
            val two = Metric(5, "custom")
            println(one.total)
            println(one.tag)
            println(two.total)
            println(two.tag)
        }
    "#,
    );
    assert_eq!(out, &["3", "base", "5", "custom"]);
}

#[test]
fn test_secondary_constructor_with_array_destructure_input() {
    let out = run_prints(
        r#"
        class Matrix {
            val rows: Int
            val cols: Int

            constructor(values: Array<Array<Int>>) {
                this.rows = values.size
                this.cols = if (values.isNotEmpty()) values[0].size else 0
            }

            constructor(rows: Int, cols: Int) : this(Array(rows) { Array(cols) { 0 } })
        }

        fun main() {
            val a = Matrix(2, 3)
            val b = Matrix(arrayOf(arrayOf(1, 2), arrayOf(3, 4), arrayOf(5, 6)))
            println(a.rows)
            println(a.cols)
            println(b.rows)
            println(b.cols)
        }
    "#,
    );
    assert_eq!(out, &["2", "3", "3", "2"]);
}

#[test]
fn test_secondary_constructor_private_validation() {
    let out = run_prints(
        r#"
        class PositiveCounter {
            val value: Int

            private constructor(value: Int) {
                this.value = value
            }

            constructor(raw: Int, valid: Boolean) : this(if (valid && raw > 0) raw else 0) {
                println("built")
            }
        }

        fun main() {
            println(PositiveCounter(-3, true).value)
            println(PositiveCounter(5, false).value)
            println(PositiveCounter(5, true).value)
        }
    "#,
    );
    assert_eq!(out, &["built", "built", "built", "0", "0", "5"]);
}

#[test]
fn test_secondary_constructor_delegation_argument_evaluation_order() {
    let out = run_prints(
        r#"
        var order = ""

        fun mark(value: String): Int {
            order += value
            return if (value == "left") 1 else 2
        }

        class Probe {
            val first: Int
            val second: Int

            constructor(first: Int, second: Int) {
                this.first = first
                this.second = second
            }

            constructor(value: Int) : this(mark("left"), mark("right") + value) {
                // body intentionally empty
            }
        }

        fun main() {
            val probe = Probe(3)
            println(order)
            println(probe.first)
            println(probe.second)
        }
    "#,
    );
    assert_eq!(out, &["leftright", "1", "5"]);
}

#[test]
fn test_secondary_constructor_vararg_forwarding() {
    let out = run_prints(
        r#"
        class VarArgProbe {
            val text: String

            constructor(prefix: String, vararg values: Int) {
                this.text = prefix + values.joinToString(":")
            }

            constructor(value: Int) : this("n", value, value + 1) {}
        }

        fun main() {
            println(VarArgProbe("v", 1, 2, 3).text)
            println(VarArgProbe(4).text)
        }
    "#,
    );
    assert_eq!(out, &["v:1:2:3", "n:4:5"]);
}

#[test]
fn test_secondary_constructor_body_executes_after_delegation_chain() {
    let out = run_prints(
        r#"
        var trace = ""

        class Chain {
            var value: Int

            constructor() {
                trace += "root;"
                value = 1
            }

            constructor(value: Int) : this() {
                trace += "inner;"
                this.value = value
            }

            constructor(value: Int, extra: Int) : this(value) {
                trace += "leaf;"
                this.value = value + extra
            }
        }

        fun main() {
            val c = Chain(2, 3)
            println(trace)
            println(c.value)
        }
    "#,
    );
    assert_eq!(out, &["root;inner;leaf;", "5"]);
}

#[test]
fn test_secondary_constructor_super_argument_is_evaluated_for_derived_init() {
    let out = run_prints(
        r#"
        var calls = ""

        open class Parent(val seed: Int) {
            val value = seed + 1
        }

        fun computeSeed(base: Int): Int {
            calls += base.toString()
            return base * 10
        }

        class Child : Parent {
            val offset: Int

            constructor(base: Int) : super(computeSeed(base)) {
                this.offset = base + 1
            }
        }

        fun main() {
            val c = Child(7)
            println(c.value)
            println(c.offset)
            println(calls)
        }
    "#,
    );
    assert_eq!(out, &["71", "8", "7"]);
}

#[test]
fn test_secondary_constructor_constructor_can_throw_and_fail() {
    let out = run_prints(
        r#"
        class Guard {
            val value: Int

            constructor(value: Int) {
                if (value < 0) {
                    throw Exception("invalid")
                }
                this.value = value
            }
        }

        fun main() {
            try {
                println(Guard(-2).value)
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#,
    );
    assert_eq!(out, &["caught"]);
}

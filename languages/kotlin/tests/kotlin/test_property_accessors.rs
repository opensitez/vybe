use crate::helpers::run_prints;

#[test]
fn test_property_getter_computed() {
    let out = run_prints(
        r#"
        class Box(val a: Int, val b: Int) {
            val sum: Int get() = a + b
        }
        fun main() {
            println(Box(2, 3).sum)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_property_setter_basic() {
    let out = run_prints(
        r#"
        class Box {
            var v: Int = 0
                set(value) { field = value }
        }
        fun main() {
            val b = Box()
            b.v = 4
            println(b.v)
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_property_setter_normalized() {
    let out = run_prints(
        r#"
        class Box {
            var value: Int = 0
                set(value) { field = if (value < 0) 0 else value }
        }
        fun main() {
            val b = Box()
            b.value = -2
            println(b.value)
            b.value = 7
            println(b.value)
        }
    "#,
    );
    assert_eq!(out, &["0", "7"]);
}

#[test]
fn test_property_getter_with_side_effect_count() {
    let out = run_prints(
        r#"
        class Box {
            var c = 0
            val label: Int
                get() {
                    c += 1
                    return c
                }
        }
        fun main() {
            val b = Box()
            println(b.label)
            println(b.label)
        }
    "#,
    );
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_property_private_setter() {
    let out = run_prints(
        r#"
        class Box {
            var value: Int = 1
                private set
        }
        fun main() {
            val b = Box()
            println(b.value)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_property_backing_field_private_setter() {
    let out = run_prints(
        r#"
        class Box {
            private var _v = 0
            var value: Int
                get() = _v
                private set
        }
        fun main() {
            val b = Box()
            println(b.value)
        }
    "#,
    );
    assert_eq!(out, &["0"]);
}

#[test]
fn test_property_delayed_init() {
    let out = run_prints(
        r#"
        class Box {
            val v: Int by lazy { 5 }
        }
        fun main() {
            val b = Box()
            println(b.v)
            println(b.v)
        }
    "#,
    );
    assert_eq!(out, &["5", "5"]);
}

#[test]
fn test_property_custom_getter_lower() {
    let out = run_prints(
        r#"
        class Point(val x: Int, val y: Int) {
            val min: Int get() = if (x < y) x else y
        }
        fun main() {
            println(Point(2, 7).min)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_property_lateinit_var() {
    let out = run_prints(
        r#"
        class Holder {
            lateinit var text: String
            fun run() {
                text = "ok"
                println(text)
            }
        }
        fun main() {
            Holder().run()
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_property_with_setter_clamp_positive() {
    let out = run_prints(
        r#"
        class Clamp {
            var value: Int = 0
                set(v) {
                    field = if (v < 0) 0 else v
                }
        }
        fun main() {
            val c = Clamp()
            c.value = -1
            println(c.value)
        }
    "#,
    );
    assert_eq!(out, &["0"]);
}

#[test]
fn test_property_getter_boolean_logic() {
    let out = run_prints(
        r#"
        class Status(val active: Boolean) {
            val activeText: String get() = if (active) "yes" else "no"
        }
        fun main() {
            println(Status(true).activeText)
            println(Status(false).activeText)
        }
    "#,
    );
    assert_eq!(out, &["yes", "no"]);
}

#[test]
fn test_property_custom_setter_chain() {
    let out = run_prints(
        r#"
        class Box {
            var value: Int = 1
                set(v) {
                    field = v
                    println("set")
                }
        }
        fun main() {
            val b = Box()
            b.value = 3
        }
    "#,
    );
    assert_eq!(out, &["set"]);
}

#[test]
fn test_property_getter_computed_length() {
    let out = run_prints(
        r#"
        class Text {
            var value: String = ""
                set(v) { field = v }
            val length: Int get() = value.length
        }
        fun main() {
            val t = Text()
            t.value = "abc"
            println(t.length)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_property_setter_with_guard() {
    let out = run_prints(
        r#"
        class Box {
            var value: Int = 0
                set(v) {
                    field = if (v == 13) 0 else v
                }
        }
        fun main() {
            val b = Box()
            b.value = 13
            println(b.value)
        }
    "#,
    );
    assert_eq!(out, &["0"]);
}

#[test]
fn test_property_getter_casting() {
    let out = run_prints(
        r#"
        class Holder {
            val number: Number = 2
            val intValue: Int get() = number.toInt()
        }
        fun main() {
            println(Holder().intValue)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_property_setter_and_getter_in_class_hierarchy() {
    let out = run_prints(
        r#"
        open class Base {
            open var value: Int = 1
        }
        class Child : Base() {
            override var value: Int = 2
        }
        fun main() {
            val b: Base = Child()
            println(b.value)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_property_backed_by_computed_expression() {
    let out = run_prints(
        r#"
        class Box {
            var a: Int = 2
            var b: Int = 3
            val sum: Int
                get() = a + b
        }
        fun main() {
            val b = Box()
            println(b.sum)
            b.a = 5
            println(b.sum)
        }
    "#,
    );
    assert_eq!(out, &["5", "8"]);
}

#[test]
fn test_property_custom_getter_transform() {
    let out = run_prints(
        r#"
        class Box(val raw: String) {
            val trimmed: String get() = raw.trim()
        }
        fun main() {
            println(Box("  x ").trimmed)
        }
    "#,
    );
    assert_eq!(out, &["x"]);
}

#[test]
fn test_property_lazy_reused() {
    let out = run_prints(
        r#"
        class Holder {
            var initCount = 0
            val data: Int by lazy {
                initCount += 1
                9
            }
        }
        fun main() {
            val h = Holder()
            println(h.initCount)
            println(h.data)
            println(h.data)
            println(h.initCount)
        }
    "#,
    );
    assert_eq!(out, &["0", "9", "9", "1"]);
}

#[test]
fn test_property_getter_calls_method() {
    let out = run_prints(
        r#"
        class Box {
            val a = 1
            val b = 2
            val total: Int get() = sum()
            fun sum() = a + b
        }
        fun main() {
            println(Box().total)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_property_in_data_class_default_getter() {
    let out = run_prints(
        r#"
        data class Point(val x: Int, val y: Int)
        fun main() {
            val p = Point(1, 2)
            println(p.x)
            println(p.y)
        }
    "#,
    );
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_property_setter_chain_validation() {
    let out = run_prints(
        r#"
        class Box {
            var value: Int = 1
                set(v) {
                    field = v
                }
        }
        fun main() {
            val b = Box()
            b.value = 4
            b.value = b.value + 1
            println(b.value)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_property_array_accessor() {
    let out = run_prints(
        r#"
        class Bag {
            private val values = IntArray(3)
            var second: Int
                get() = values[1]
                set(v) { values[1] = v }
        }
        fun main() {
            val b = Bag()
            b.second = 7
            println(b.second)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_property_getter_nullable() {
    let out = run_prints(
        r#"
        class Box {
            var text: String? = null
            val safe: String get() = text ?: "none"
        }
        fun main() {
            println(Box().safe)
            val b = Box()
            b.text = "x"
            println(b.safe)
        }
    "#,
    );
    assert_eq!(out, &["none", "x"]);
}

#[test]
fn test_property_visibility_private_field() {
    let out = run_prints(
        r#"
        class Box {
            private var _v = 1
            var v: Int
                get() = _v
                set(value) { _v = value }
        }
        fun main() {
            val b = Box()
            b.v = 12
            println(b.v)
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_property_double_access_in_loop() {
    let out = run_prints(
        r#"
        class Count {
            var n = 0
                private set
            fun inc() { n += 1 }
        }
        fun main() {
            val c = Count()
            c.inc(); c.inc()
            println(c.n)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_property_with_custom_equals() {
    let out = run_prints(
        r#"
        class Box(val value: Int) {
            val isPositive: Boolean get() = value > 0
        }
        fun main() {
            val a = Box(3)
            val b = Box(-1)
            println(a.isPositive)
            println(b.isPositive)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_property_setter_side_effect_log() {
    let out = run_prints(
        r#"
        class Logger {
            var value: Int = 0
                set(v) {
                    println(v)
                    field = v
                }
        }
        fun main() {
            val l = Logger()
            l.value = 1
            l.value = 2
        }
    "#,
    );
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_property_nested_in_getter() {
    let out = run_prints(
        r#"
        class Box {
            var v = 1
            val double: Int get() = run {
                val m = v * 2
                m
            }
        }
        fun main() {
            println(Box().double)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_property_delegates_to_method() {
    let out = run_prints(
        r#"
        class Box {
            var value: Int = 3
            val visible: Int get() = display()
            private fun display() = value
        }
        fun main() {
            println(Box().visible)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_property_custom_interface_implementation() {
    let out = run_prints(
        r#"
        interface HasValue { var value: Int }
        class Box(override var value: Int) : HasValue
        fun main() {
            val b = Box(8)
            println(b.value)
            b.value = 9
            println(b.value)
        }
    "#,
    );
    assert_eq!(out, &["8", "9"]);
}

#[test]
fn test_property_getter_with_range_check() {
    let out = run_prints(
        r#"
        class Counter {
            var value = 0
            val isSmall: Boolean get() = value in 0..10
        }
        fun main() {
            val c = Counter()
            println(c.isSmall)
            c.value = 15
            println(c.isSmall)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

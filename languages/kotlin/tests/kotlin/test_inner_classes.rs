kotlin_run_test!(
    test_inner_class_reads_outer_property,
    r#"
        class Outer(val base: Int) {
            inner class Inner(val delta: Int)
            fun make(): Int = Inner(3).delta + base
        }

        fun main() {
            println(Outer(10).make())
        }
    "#,
    &["13"]
);

kotlin_run_test!(
    test_inner_class_method_calls_outer,
    r#"
        class Counter(val prefix: String) {
            inner class Marker {
                fun label(v: Int): String = "$prefix-$v"
            }
        }

        fun main() {
            println(Counter("k").Marker().label(9))
        }
    "#,
    &["k-9"]
);

kotlin_run_test!(
    test_inner_class_accesses_outer_mutable_state,
    r#"
        class Store {
            var value = 1
            inner class Bump {
                fun add(v: Int) {
                    value += v
                }
            }
        }

        fun main() {
            val store = Store()
            val bump = store.Bump()
            bump.add(5)
            println(store.value)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_inner_class_from_method,
    r#"
        class Builder {
            private val base = 2
            inner class Worker(val factor: Int) {
                fun total(): Int = base * factor
            }

            fun make(): Int = Worker(4).total()
        }

        fun main() {
            println(Builder().make())
        }
    "#,
    &["8"]
);

kotlin_run_test!(
    test_multiple_inner_instances_share_outer,
    r#"
        class Tracker {
            var ticks = 0
            inner class Probe {
                fun hit() { ticks += 1 }
            }
        }

        fun main() {
            val t = Tracker()
            val a = t.Probe()
            val b = t.Probe()
            a.hit()
            b.hit()
            println(t.ticks)
        }
    "#,
    &["2"]
);

kotlin_run_test!(
    test_inner_class_chain,
    r#"
        class Network {
            val root = "R"
            inner class Segment {
                inner class Node(val id: Int) {
                    fun label(): String = root + id
                }
            }
        }

        fun main() {
            val node = Network().Segment().Node(7)
            println(node.label())
        }
    "#,
    &["R7"]
);

kotlin_run_test!(
    test_inner_class_returning_outer_reference,
    r#"
        class Logger {
            private var tag = "log"
            inner class Entry {
                fun marker(): Logger = this@Logger
            }

            fun tag(): String = tag
        }

        fun main() {
            val entry = Logger().Entry()
            println(entry.marker().tag())
        }
    "#,
    &["log"]
);

kotlin_run_test!(
    test_inner_class_with_extension_style_api,
    r#"
        class Session {
            val prefix = "S"
            inner class Formatter {
                fun apply(v: String): String = prefix + v
            }
        }

        fun main() {
            val out = Session().Formatter().apply("tep")
            println(out)
        }
    "#,
    &["Step"]
);

kotlin_run_test!(
    test_inner_class_overrides_outer_property_after_change,
    r#"
        class Counter {
            var base = 1
            inner class Ticker {
                fun tick() { base += 2 }
            }

            fun value(): Int = base
        }

        fun main() {
            val c = Counter()
            val t = c.Ticker()
            t.tick()
            t.tick()
            println(c.value())
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_inner_class_in_nested_generic_context,
    r#"
        class Generic<T>(val value: T) {
            inner class Holder {
                fun asString(): String = value.toString()
            }
        }

        fun main() {
            println(Generic(12).Holder().asString())
        }
    "#,
    &["12"]
);

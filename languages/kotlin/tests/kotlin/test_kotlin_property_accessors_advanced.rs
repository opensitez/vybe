kotlin_run_test!(
    test_getter_derived_value,
    r#"
        class Holder {
            private val value = 4
            val doubled get() = value * 2
        }

        fun main() {
            println(Holder().doubled)
        }
    "#,
    &["8"]
);

kotlin_run_test!(
    test_setter_validation_with_private_backing,
    r#"
        class Counter {
            private var _count = 0
            var count: Int
                get() = _count
                set(v) {\n                    _count = if (v < 0) 0 else v
                }
        }

        fun main() {
            val c = Counter()
            c.count = -5
            println(c.count)
            c.count = 7
            println(c.count)
        }
    "#,
    &["0", "7"]
);

kotlin_run_test!(
    test_custom_accessor_computes_transient_state,
    r#"
        class Meter {
            private var ticks = 0
            var total: Int
                get() = ticks + 1
                set(v) { ticks = v }
        }

        fun main() {
            val m = Meter()
            m.total = 3
            println(m.total)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_backing_field_visible_to_accessors,
    r#"
        class Name {
            var text: String = "x"
                set(value) {
                    field = value.uppercase()
                }
        }

        fun main() {
            val n = Name()
            n.text = "ab"
            println(n.text)
        }
    "#,
    &["AB"]
);

kotlin_run_test!(
    test_lazy_property_initializer,
    r#"
        class Expensive {
            val value by lazy {
                3 + 4
            }
        }

        fun main() {
            val e = Expensive()
            println(e.value)
            println(e.value)
        }
    "#,
    &["7", "7"]
);

kotlin_run_test!(
    test_accessor_visibility_modifiers,
    r#"
        class Item {
            var value: Int = 1
                private set
                public get
            fun add(v: Int) { value += v }
        }

        fun main() {
            val i = Item()
            i.add(2)
            println(i.value)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_custom_indexing_like_property,
    r#"
        class Bag {
            private val data = listOf(1, 3, 5)
            operator fun get(index: Int): Int = data[index]
            val head get() = data.first()
        }

        fun main() {
            val b = Bag()
            println(b[2])
            println(b.head)
        }
    "#,
    &["5", "1"]
);

kotlin_run_test!(
    test_setter_with_previous_state,
    r#"
        class Running {
            private var _sum = 0
            var sum: Int
                get() = _sum
                set(value) { _sum += value }
        }

        fun main() {
            val r = Running()
            r.sum = 3
            r.sum = 4
            println(r.sum)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_property_with_compute_side_effect,
    r#"
        class Toggle {
            private var on = false
            val flag: Boolean
                get() {
                    on = !on
                    return on
                }
        }

        fun main() {
            val t = Toggle()
            println(t.flag)
            println(t.flag)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_property_reassigns_and_getter_validation,
    r#"
        class RangeValue {
            private var _value = 0
            var value: Int
                get() = _value
                set(v) { _value = if (v > 10) 10 else v }
        }

        fun main() {
            val r = RangeValue()
            r.value = 15
            println(r.value)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_lateinit_access_before_initialization_throws,
    r#"
        class Box {
            lateinit var text: String
        }

        fun main() {
            val box = Box()
            val result = try {
                println(box.text)
                "ok"
            } catch (e: UninitializedPropertyAccessException) {
                "uninitialized"
            }
            println(result)
        }
    "#,
    &["uninitialized"]
);

kotlin_run_test!(
    test_lateinit_can_be_set_and_read,
    r#"
        class Box {
            lateinit var text: String
        }

        fun main() {
            val box = Box()
            box.text = "k"
            println(box.text)
        }
    "#,
    &["k"]
);

kotlin_run_test!(
    test_lateinit_assigned_in_method,
    r#"
        class Container {
            lateinit var value: String

            fun prepare() {
                value = "ready"
            }
        }

        fun main() {
            val c = Container()
            c.prepare()
            println(c.value)
        }
    "#,
    &["ready"]
);

kotlin_run_test!(
    test_lateinit_with_reassignment,
    r#"
        class Bag {
            lateinit var label: String

            fun setA() { label = "a" }
            fun setB() { label = "b" }
        }

        fun main() {
            val b = Bag()
            b.setA()
            println(b.label)
            b.setB()
            println(b.label)
        }
    "#,
    &["a", "b"]
);

kotlin_run_test!(
    test_lateinit_in_multiple_instances,
    r#"
        class Holder {
            lateinit var note: String
        }

        fun main() {
            val a = Holder()
            val b = Holder()
            a.note = "A"
            b.note = "B"
            println(a.note)
            println(b.note)
        }
    "#,
    &["A", "B"]
);

kotlin_run_test!(
    test_lateinit_with_list_type,
    r#"
        class Collector {
            lateinit var values: MutableList<Int>
        }

        fun main() {
            val c = Collector()
            c.values = mutableListOf(1, 2)
            c.values.add(3)
            println(c.values.joinToString(","))
        }
    "#,
    &["1,2,3"]
);

kotlin_run_test!(
    test_lateinit_in_constructor_secondary,
    r#"
        class Payload {
            lateinit var text: String

            constructor(v: String) {
                text = v
            }
        }

        fun main() {
            val p = Payload("go")
            println(p.text)
        }
    "#,
    &["go"]
);

kotlin_run_test!(
    test_lateinit_with_try_after_set,
    r#"
        class Probe {
            lateinit var text: String
        }

        fun main() {
            val p = Probe()
            p.text = "x"
            val result = try {
                p.text
                "after-set"
            } catch (e: Exception) {
                "bad"
            }
            println(result)
        }
    "#,
    &["after-set"]
);

kotlin_run_test!(
    test_lateinit_property_reassigned_after_throw,
    r#"
        class Holder {
            lateinit var value: String

            fun first() { value = "v1" }
            fun second() { value = "v2" }
        }

        fun main() {
            val h = Holder()
            val a = try {
                h.value
                "ok"
            } catch (e: UninitializedPropertyAccessException) {
                "first-failed"
            }
            h.first()
            println(a)
            h.second()
            println(h.value)
        }
    "#,
    &["first-failed", "v2"]
);

kotlin_run_test!(
    test_lateinit_with_function_call_dependency,
    r#"
        class Source {
            lateinit var source: String
        }

        fun build(s: Source): String = s.source

        fun main() {
            val s = Source()
            s.source = "done"
            println(build(s))
        }
    "#,
    &["done"]
);

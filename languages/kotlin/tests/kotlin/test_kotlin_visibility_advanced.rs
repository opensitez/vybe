kotlin_run_test!(
    test_protected_member_visible_in_subclass,
    r#"
        open class Base {
            protected fun token() = "ok"
        }

        class Child : Base() {
            fun reveal() = token()
        }

        fun main() {
            println(Child().reveal())
        }
    "#,
    &["ok"]
);

kotlin_run_test!(
    test_private_setter_blocks_external_mutation,
    r#"
        class Counter {
            var value: Int = 0
                private set

            fun bump() { value += 1 }
        }

        fun main() {
            val counter = Counter()
            counter.bump()
            counter.bump()
            println(counter.value)
        }
    "#,
    &["2"]
);

kotlin_run_test!(
    test_internal_visibility_within_module,
    r#"
        internal class Box {
            fun payload() = 7
        }

        fun main() {
            println(Box().payload())
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_private_top_level_function_stays_in_file,
    r#"
        private fun secret(): String = "hidden"

        fun main() {
            println(secret())
        }
    "#,
    &["hidden"]
);

kotlin_run_test!(
    test_public_open_override_is_visible_via_reference,
    r#"
        open class Parent { open val name: String = "p" }
        class Child : Parent() { override val name: String = "c" }

        fun main() {
            val base: Parent = Child()
            println(base.name)
        }
    "#,
    &["c"]
);

kotlin_run_test!(
    test_local_visibility_on_constructor_property,
    r#"
        class Holder(private val tag: String) {
            fun render() = tag
        }

        fun main() {
            println(Holder("x").render())
        }
    "#,
    &["x"]
);

kotlin_run_test!(
    test_property_visibility_in_data_holder,
    r#"
        class Repo {
            private val raw = 1
            val exposed get() = raw
        }

        fun main() {
            println(Repo().exposed)
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_private_setter_with_external_caller,
    r#"
        class Bucket {
            var total = 0
                private set

            fun add(v: Int) {
                total += v
            }
        }

        fun main() {
            val b = Bucket()
            b.add(3)
            b.add(4)
            println(b.total)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_abstract_visibility_with_override_chain,
    r#"
        abstract class Root {
            protected abstract fun token(): String
        }

        class Leaf : Root() {
            override fun token() = "seen"
        }

        fun main() {
            println(Leaf().token())
        }
    "#,
    &["seen"]
);

kotlin_run_test!(
    test_nested_class_visibility,
    r#"
        class Outer {
            private val secret = "open"
            inner class Inner {
                fun reveal() = secret
            }
        }

        fun main() {
            println(Outer().Inner().reveal())
        }
    "#,
    &["open"]
);

kotlin_run_test!(
    test_visibility_of_extension_target,
    r#"
        private class Core {
            fun value() = 9
        }

        fun Core.expose() = value()

        fun main() {
            println(Core().expose())
        }
    "#,
    &["9"]
);

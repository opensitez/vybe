kotlin_run_test!(
    test_primary_constructor_and_init_order,
    r#"
        class Counter start {
            val value: Int = start
            init {
                println(value)
            }
            constructor(): this(5)
        }

        fun main() {
            Counter()
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_secondary_constructor_chains,
    r#"
        class Node {
            val x: Int
            constructor(a: Int) {
                x = a
            }
            constructor(a: Int, b: Int) : this(a + b)
        }

        fun main() {
            val n = Node(2, 3)
            println(n.x)
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_property_initialized_from_constructor_parameter_order,
    r#"
        class User(name: String) {
            val upper = name.uppercase()
            init { println(upper) }
            val size = name.length
        }

        fun main() {
            println(User("ab").size)
        }
    "#,
    &["AB", "2"]
);

kotlin_run_test!(
    test_init_blocks_run_once_per_instance,
    r#"
        class Box {
            init { println(1) }
            init { println(2) }
        }

        fun main() {
            Box()
            Box()
        }
    "#,
    &["1", "2", "1", "2"]
);

kotlin_run_test!(
    test_constructor_with_optional_default,
    r#"
        class Entry(val a: Int = 1, val b: Int = 2) {
            val sum = a + b
        }

        fun main() {
            println(Entry(3).sum)
            println(Entry().sum)
        }
    "#,
    &["5", "3"]
);

kotlin_run_test!(
    test_init_can_reference_previous_properties,
    r#"
        class Product {
            val base = 3
            val doubled = base * 2
            init {
                println(doubled)
            }
        }

        fun main() {
            println(Product().base)
        }
    "#,
    &["6", "3"]
);

kotlin_run_test!(
    test_secondary_constructor_with_default_chain,
    r#"
        class Config {
            val mode: String
            constructor() { mode = "auto" }
            constructor(raw: String) : this() { mode = raw }
        }

        fun main() {
            println(Config().mode)
            println(Config("manual").mode)
        }
    "#,
    &["auto", "manual"]
);

kotlin_run_test!(
    test_init_block_updates_mutable_state,
    r#"
        class Tracker {
            var count = 0
            init { count = count + 1 }
        }

        fun main() {
            val t = Tracker()
            println(t.count)
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_companion_object_called_from_constructor,
    r#"
        class Id {
            val value: Int
            init {
                value = next()
            }

            companion object {
                private var seq = 0
                fun next(): Int {
                    seq += 1
                    return seq
                }
            }
        }

        fun main() {
            println(Id().value)
            println(Id().value)
        }
    "#,
    &["1", "2"]
);

kotlin_run_test!(
    test_inner_constructor_for_data_properties,
    r#"
        class Profile {
            val name: String
            val suffix: String
            constructor(name: String) {
                this.name = name
                this.suffix = name.takeLast(1)
            }
            constructor(name: String, idx: Int) : this(name) {
                this.suffix = name[idx]
            }

            fun render(): String = name + suffix
        }

        fun main() {
            println(Profile("abc").render())
            println(Profile("xyz", 1).render())
        }
    "#,
    &["abcc", "xyy"]
);

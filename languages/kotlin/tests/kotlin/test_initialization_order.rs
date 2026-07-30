use crate::helpers::run_prints;

#[test]
fn test_primary_constructor_property_initializes_before_init_block() {
    let out = run_prints(r#"
        class Holder(val base: Int) {
            val plus = base + 1
            val label: String

            init {
                label = "v=" + plus.toString()
            }

            fun out(): String = label
        }

        fun main() {
            println(Holder(3).out())
        }
    "#);
    assert_eq!(out, &["v=4"]);
}

#[test]
fn test_init_blocks_execute_in_top_down_order() {
    let out = run_prints(r#"
        class Base {
            init {
                println("base")
            }
        }

        class Leaf : Base() {
            init {
                println("leaf")
            }
        }

        fun main() {
            Leaf()
        }
    "#);
    assert_eq!(out, &["base", "leaf"]);
}

#[test]
fn test_property_init_order_with_override_chain() {
    let out = run_prints(r#"
        open class Base {
            open val base = 1
            init {
                println(base)
            }
        }

        class Child : Base() {
            override val base = 4
            init {
                println(base)
            }
        }

        fun main() {
            Child()
        }
    "#);
    assert_eq!(out, &["4", "4"]);
}

#[test]
fn test_initialization_order_does_not_recompute_dependencies() {
    let out = run_prints(r#"
        var ticks = 0

        class Holder {
            val first = next()
            val second = first + 1

            init {
                println(first)
                println(second)
            }
        }

        fun next(): Int {
            ticks += 1
            return ticks
        }

        fun main() {
            Holder()
            println(ticks)
        }
    "#);
    assert_eq!(out, &["1", "2", "1"]);
}

#[test]
fn test_init_evaluates_properties_before_secondary_constructor() {
    let out = run_prints(r#"
        class Holder(val value: Int) {
            val label: String

            init {
                label = "v=" + value.toString()
            }

            constructor() : this(5)

            fun out(): String = label
        }

        fun main() {
            val item = Holder()
            println(item.out())
        }
    "#);
    assert_eq!(out, &["v=5"]);
}

#[test]
fn test_init_block_can_reference_secondary_defaults() {
    let out = run_prints(r#"
        class Holder {
            val value: Int

            init {
                value = 7
                println("init")
            }

            constructor() {
                this()
            }

            fun out(): Int = value
        }

        fun main() {
            val item = Holder()
            println(item.out())
        }
    "#);
    assert_eq!(out, &["init", "7"]);
}

#[test]
fn test_initialization_logs_in_nested_property_chain() {
    let out = run_prints(r#"
        class Holder {
            val first = 1
            val second = first + one()
            val third = second + 1

            init {
                println(third)
            }
        }

        fun one(): Int = 2

        fun main() {
            Holder()
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_companion_and_init_are_independent_timelines() {
    let out = run_prints(r#"
        class Holder {
            companion object {
                init { println("companion") }
            }

            init {
                println("instance")
            }
        }

        fun main() {
            Holder()
            Holder()
        }
    "#);
    assert_eq!(out, &["companion", "instance", "instance"]);
}

#[test]
fn test_initialization_of_multiple_properties_order_by_appearance() {
    let out = run_prints(r#"
        class Holder {
            val a = 1
            val b = a + c
            val c = 3

            init {
                println(a)
                println(b)
                println(c)
            }
        }

        fun main() {
            Holder()
        }
    "#);
    assert_eq!(out, &["1", "4", "3"]);
}

#[test]
fn test_init_without_secondary_constructor_still_runs_defaults() {
    let out = run_prints(r#"
        class Holder(val base: Int = 1) {
            val scaled = base * 2
            init { println(scaled) }
        }

        fun main() {
            Holder()
            Holder(4)
        }
    "#);
    assert_eq!(out, &["2", "8"]);
}

#[test]
fn test_init_of_derived_uses_base_initialized_state() {
    let out = run_prints(r#"
        open class Base {
            val base = 10
            init { println(base) }
        }

        class Child : Base() {
            val child = base + 1
            init { println(child) }
        }

        fun main() {
            Child()
        }
    "#);
    assert_eq!(out, &["10", "11"]);
}

#[test]
fn test_initialization_of_local_class_occurs_on_use() {
    let out = run_prints(r#"
        fun main() {
            class Local {
                init { println("local") }
            }

            println("start")
            Local()
            println("end")
        }
    "#);
    assert_eq!(out, &["start", "local", "end"]);
}

use crate::helpers::run_prints;

#[test]
fn test_primary_constructor_property_initializes_before_init_block() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["v=4"]);
}

#[test]
fn test_init_blocks_execute_in_top_down_order() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["base", "leaf"]);
}

#[test]
fn test_property_init_order_with_override_chain() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["4", "4"]);
}

#[test]
fn test_initialization_order_does_not_recompute_dependencies() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["1", "2", "1"]);
}

#[test]
fn test_init_evaluates_properties_before_secondary_constructor() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["v=5"]);
}

#[test]
fn test_init_block_can_reference_secondary_defaults() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["init", "7"]);
}

#[test]
fn test_initialization_logs_in_nested_property_chain() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_companion_and_init_are_independent_timelines() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["companion", "instance", "instance"]);
}

#[test]
fn test_initialization_of_multiple_properties_order_by_appearance() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["1", "4", "3"]);
}

#[test]
fn test_init_without_secondary_constructor_still_runs_defaults() {
    let out = run_prints(
        r#"
        class Holder(val base: Int = 1) {
            val scaled = base * 2
            init { println(scaled) }
        }

        fun main() {
            Holder()
            Holder(4)
        }
    "#,
    );
    assert_eq!(out, &["2", "8"]);
}

#[test]
fn test_init_of_derived_uses_base_initialized_state() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["10", "11"]);
}

#[test]
fn test_initialization_of_local_class_occurs_on_use() {
    let out = run_prints(
        r#"
        fun main() {
            class Local {
                init { println("local") }
            }

            println("start")
            Local()
            println("end")
        }
    "#,
    );
    assert_eq!(out, &["start", "local", "end"]);
}

#[test]
fn test_init_order_records_property_evaluation_and_init() {
    let out = run_prints(
        r#"
        var trace = ""
        fun tick(value: String): Int {
            trace += value
            return value.length
        }

        class Holder {
            val first = tick("a")
            val second = tick("b")
            init {
                println(trace)
            }
            val third = first + second
            init {
                println(third)
            }
        }

        fun main() {
            Holder()
        }
    "#,
    );
    assert_eq!(out, &["ab", "2"]);
}

#[test]
fn test_secondary_constructor_delegation_preserves_base_initialization() {
    let out = run_prints(
        r#"
        open class Base(prefix: Int) {
            val value = prefix
            init {
                println(value)
            }
        }

        class Leaf : Base {
            val label: Int
            constructor() : this(2)
            constructor(seed: Int) : super(seed) {
                label = seed * 10
            }
        }

        fun main() {
            println(Leaf().label)
        }
    "#,
    );
    assert_eq!(out, &["2", "20"]);
}

#[test]
fn test_derived_class_init_runs_after_base_init() {
    let out = run_prints(
        r#"
        open class Base {
            init {
                println("base")
            }
        }

        class Child : Base() {
            init {
                println("child-1")
            }

            init {
                println("child-2")
            }
        }

        fun main() {
            Child()
        }
    "#,
    );
    assert_eq!(out, &["base", "child-1", "child-2"]);
}

#[test]
fn test_field_initialization_happens_in_declaration_order() {
    let out = run_prints(
        r#"
        class Holder {
            val first = 1
            val second = first + 1
            val third = second + first
            init {
                println(first)
                println(second)
                println(third)
            }
        }

        fun main() {
            Holder()
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "3"]);
}

#[test]
fn test_companion_init_occurs_before_instance_init_once() {
    let out = run_prints(
        r#"
        class Holder {
            companion object {
                var count = 0
                init {
                    count = 9
                }
            }

            init {
                println(companionInit())
            }
        }

        fun companionInit(): Int = Holder.count

        fun main() {
            Holder()
            Holder()
        }
    "#,
    );
    assert_eq!(out, &["9", "9"]);
}

#[test]
fn test_init_uses_primary_constructor_parameters() {
    let out = run_prints(
        r#"
        class Holder(prefix: String) {
            val value: String

            init {
                value = prefix.uppercase()
                println(value)
            }
        }

        fun main() {
            Holder("x")
            Holder("y")
        }
    "#,
    );
    assert_eq!(out, &["X", "Y"]);
}

#[test]
fn test_init_block_can_read_property_updates() {
    let out = run_prints(
        r#"
        var factor = 1

        class Holder {
            val value = factor

            init {
                factor = 4
            }

            val adjusted = value * factor

            init {
                println(value)
                println(adjusted)
            }
        }

        fun main() {
            Holder()
            println(factor)
        }
    "#,
    );
    assert_eq!(out, &["1", "4", "4"]);
}

#[test]
fn test_init_block_for_local_class_runs_at_instantiation() {
    let out = run_prints(
        r#"
        fun main() {
            class Local {
                init { println("local-init") }
            }

            println("a")
            Local()
            println("b")
        }
    "#,
    );
    assert_eq!(out, &["a", "local-init", "b"]);
}

#[test]
fn test_inherited_property_shadowing_does_not_reorder_init() {
    let out = run_prints(
        r#"
        open class Base {
            open val value = 2
        }

        class Child : Base() {
            override val value = 7
            val total = value + 1

            init {
                println(value)
                println(total)
            }
        }

        fun main() {
            Child()
        }
    "#,
    );
    assert_eq!(out, &["7", "8"]);
}

#[test]
fn test_init_blocks_in_multiple_levels_chain() {
    let out = run_prints(
        r#"
        open class LevelOne {
            init {
                println("one")
            }
        }

        open class LevelTwo : LevelOne() {
            init {
                println("two")
            }
        }

        class LevelThree : LevelTwo() {
            init {
                println("three")
            }
        }

        fun main() {
            LevelThree()
        }
    "#,
    );
    assert_eq!(out, &["one", "two", "three"]);
}

#[test]
fn test_constructor_default_values_apply_per_instance() {
    let out = run_prints(
        r#"
        class Holder(prefix: Int = 1) {
            val value = prefix + 1
            init {
                println(value)
            }
        }

        fun main() {
            Holder()
            Holder(2)
        }
    "#,
    );
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_init_uses_global_counter_and_reuses_updated_state() {
    let out = run_prints(
        r#"
        var stamp = 0

        fun next_stamp(): Int {
            stamp += 1
            return stamp
        }

        class Holder {
            val first = next_stamp()
            val second = first + 10
            init {
                println(first)
                println(second)
            }
            val third = next_stamp()
            init {
                println(third)
            }
        }

        fun main() {
            Holder()
            Holder()
            println(stamp)
        }
    "#,
    );
    assert_eq!(out, &["1", "11", "2", "3", "13", "4", "4"]);
}

#[test]
fn test_overridden_property_is_visible_before_derived_fields_init() {
    let out = run_prints(
        r#"
        open class Base {
            open val label: String = "base"

            init {
                println(label)
            }
        }

        class Child : Base() {
            override val label: String = "child"
            val extended = label + ":v"

            init {
                println(extended)
            }
        }

        fun main() {
            Child()
        }
    "#,
    );
    assert_eq!(out, &["child", "child:v"]);
}

#[test]
fn test_init_blocks_can_mutate_companion_state_before_subsequent_property_init() {
    let out = run_prints(
        r#"
        var globalValue = 1

        class Holder {
            val first = globalValue

            init {
                globalValue = 10
            }

            val second = globalValue * 2

            init {
                println(first)
                println(second)
            }
        }

        fun main() {
            Holder()
            Holder()
            println(globalValue)
        }
    "#,
    );
    assert_eq!(out, &["1", "20", "10", "20", "10"]);
}

#[test]
fn test_secondary_constructor_chain_initializes_once_per_instance() {
    let out = run_prints(
        r#"
        class Holder {
            val value: Int

            init {
                println("instance")
            }

            constructor() : this(3) {
                println("delegated")
            }

            constructor(seed: Int) {
                value = seed * 2
                println(value)
            }
        }

        fun main() {
            Holder()
        }
    "#,
    );
    assert_eq!(out, &["instance", "6", "delegated"]);
}

#[test]
fn test_property_initializers_evaluate_in_declaration_order() {
    let out = run_prints(
        r#"
        class Holder {
            val base = 1
            val multiplied = base * 2
            val summed = multiplied + 1

            init {
                println(base)
                println(multiplied)
                println(summed)
            }
        }

        fun main() {
            Holder()
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "3"]);
}

#[test]
fn test_local_class_initialization_runs_after_function_code_before_use() {
    let out = run_prints(
        r#"
        fun main() {
            println("pre")

            class Holder {
                val value = "init"

                init {
                    println(value)
                }
            }

            println("post")
            Holder()
            println("done")
        }
    "#,
    );
    assert_eq!(out, &["pre", "post", "init", "done"]);
}

#[test]
fn test_init_blocks_execute_in_chain_before_instance_values_are_printed() {
    let out = run_prints(
        r#"
        open class LevelOne {
            open val base = "one"
            init {
                println(base)
            }
        }

        open class LevelTwo : LevelOne() {
            init {
                println(base + "-two")
            }
        }

        class LevelThree : LevelTwo() {
            init {
                println(base + "-three")
            }
        }

        fun main() {
            LevelThree()
        }
    "#,
    );
    assert_eq!(out, &["one", "one-two", "one-three"]);
}

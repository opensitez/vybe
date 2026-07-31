use crate::helpers::run_prints;

#[test]
fn test_singleton_object_holds_mutable_state() {
    let out = run_prints(r#"
        object Counter {
            var value = 0
            fun inc() { value += 1 }
        }

        fun main() {
            Counter.inc()
            Counter.inc()
            println(Counter.value)
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_object_can_implement_interface() {
    let out = run_prints(r#"
        interface Handler {
            fun call(value: Int): Int
        }

        object PlusOne : Handler {
            override fun call(value: Int): Int = value + 1
        }

        fun apply(handler: Handler, value: Int): Int = handler.call(value)

        fun main() {
            println(apply(PlusOne, 4))
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_object_used_as_stateful_factory_target() {
    let out = run_prints(r#"
        object Factory {
            fun create(label: String): Holder = Holder(label)
        }

        class Holder(val label: String)

        fun main() {
            println(Factory.create("x").label)
        }
    "#);
    assert_eq!(out, &["x"]);
}

#[test]
fn test_nested_object_declaration() {
    let out = run_prints(r#"
        class Holder {
            object Defaults {
                val label = "ok"
            }
        }

        fun main() {
            println(Holder.Defaults.label)
        }
    "#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_object_without_state_is_reusable_singleton_reference() {
    let out = run_prints(r#"
        object Marker

        fun main() {
            println(Marker === Marker)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_object_access_in_companion_with_private_constructor() {
    let out = run_prints(r#"
        class Widget private constructor(val label: String) {
            companion object Holder {
                val shared = Maker
            }
        }

        object Maker : (String) -> Widget {
            override fun invoke(label: String): Widget = Widget(label)
        }

        fun main() {
            println(Widget.Holder.shared("x").label)
        }
    "#);
    assert_eq!(out, &["x"]);
}

#[test]
fn test_object_reference_stability_across_calls() {
    let out = run_prints(r#"
        object Counter {
            var value = 1
            fun reset() { value = 0 }
        }

        fun touch(): Int {
            Counter.value += 1
            return Counter.value
        }

        fun main() {
            Counter.reset()
            println(touch())
            println(touch())
            println(Counter.value)
        }
    "#);
    assert_eq!(out, &["1", "2", "2"]);
}

#[test]
fn test_object_properties_can_reference_companion() {
    let out = run_prints(r#"
        object Config {
            val enabled = true
        }

        class Processor {
            fun active(): String = if (Config.enabled) "yes" else "no"
        }

        fun main() {
            println(Processor().active())
        }
    "#);
    assert_eq!(out, &["yes"]);
}

#[test]
fn test_object_expression_and_object_declaration_difference() {
    let out = run_prints(r#"
        object Labeler {
            fun label(value: Int): String = "v" + value
        }

        fun main() {
            val value = object {
                fun label(value: Int): String = "local" + value
            }
            println(value.label(4))
            println(Labeler.label(4))
        }
    "#);
    assert_eq!(out, &["local4", "v4"]);
}

#[test]
fn test_object_can_hold_private_functions() {
    let out = run_prints(r#"
        object Counter {
            private fun step(value: Int): Int = value + 1
            fun next(value: Int): Int = step(value)
        }

        fun main() {
            println(Counter.next(4))
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_object_with_init_like_setup_function() {
    let out = run_prints(r#"
        object Registry {
            var value = 0
            fun setup() {
                value = 3
            }
        }

        fun main() {
            Registry.setup()
            println(Registry.value)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_object_implements_multiple_interfaces() {
    let out = run_prints(r#"
        interface Named {
            fun name(): String
        }

        interface Versioned {
            fun version(): Int
        }

        object Metadata : Named, Versioned {
            override fun name(): String = "meta"
            override fun version(): Int = 1
        }

        fun main() {
            println(Metadata.name())
            println(Metadata.version())
        }
    "#);
    assert_eq!(out, &["meta", "1"]);
}

#[test]
fn test_nested_object_members_can_be_shared_state() {
    let out = run_prints(r#"
        class Container {
            object State {
                var value = 0
            }
        }

        fun main() {
            Container.State.value += 4
            Container.State.value += 1
            println(Container.State.value)
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_object_expression_uses_local_capture() {
    let out = run_prints(r#"
        fun main() {
            val prefix = "ok"
            val value = object {
                fun label(value: Int): String = prefix + value.toString()
            }
            println(value.label(3))
        }
    "#);
    assert_eq!(out, &["ok3"]);
}

#[test]
fn test_object_expression_can_implement_custom_interface() {
    let out = run_prints(r#"
        interface Printer {
            fun print(): String
        }

        fun main() {
            val value: Printer = object : Printer {
                override fun print(): String = "done"
            }
            println(value.print())
        }
    "#);
    assert_eq!(out, &["done"]);
}

#[test]
fn test_object_reference_comparison_is_identity() {
    let out = run_prints(r#"
        object Holder {
            val value = 3
        }

        fun main() {
            println(Holder === Holder)
            println(Holder === object : Any() { val value = 3 })
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_object_as_factory_function_type() {
    let out = run_prints(r#"
        object Factory {
            fun build(prefix: String): String = prefix + "hash"
        }

        fun main() {
            println(Factory.build("a"))
            println(Factory.build("b"))
        }
    "#);
    assert_eq!(out, &["ahash", "bhash"]);
}

#[test]
fn test_object_can_be_extended_from_open_class() {
    let out = run_prints(r#"
        open class Base {
            open fun tag(): String = "base"
        }

        object Child : Base() {
            override fun tag(): String = "child"
        }

        fun main() {
            println(Child.tag())
        }
    "#);
    assert_eq!(out, &["child"]);
}

#[test]
fn test_object_can_implement_function_interface() {
    let out = run_prints(r#"
        object Incrementer : (Int) -> Int {
            override fun invoke(value: Int): Int = value + 1
        }

        fun main() {
            println(Incrementer(2))
            println(Incrementer.invoke(4))
        }
    "#);
    assert_eq!(out, &["3", "5"]);
}

#[test]
fn test_object_declaration_inside_function_has_local_scope() {
    let out = run_prints(r#"
        fun main() {
            object Local {
                val value = "local"
            }

            println(Local.value)
        }
    "#);
    assert_eq!(out, &["local"]);
}

#[test]
fn test_object_expression_returns_distinct_instance_each_call() {
    let out = run_prints(r#"
        fun builder(): Any {
            return object {
                val value = 1
            }
        }

        fun main() {
            val first = builder()
            val second = builder()
            println(first === second)
        }
    "#);
    assert_eq!(out, &["false"]);
}

#[test]
fn test_object_singleton_can_be_forwarded_through_function() {
    let out = run_prints(r#"
        object Service {
            fun id(): Int = 1
        }

        fun getService(): Any {
            return Service
        }

        fun main() {
            println(getService() === Service)
            println(getService() !== null)
            println(Service.id())
        }
    "#);
    assert_eq!(out, &["true", "true", "1"]);
}

#[test]
fn test_object_can_override_open_class_members() {
    let out = run_prints(r#"
        open class Logger {
            open fun level(): String = "base"
        }

        object Runtime : Logger() {
            override fun level(): String = "runtime"
        }

        fun main() {
            println(Runtime.level())
        }
    "#);
    assert_eq!(out, &["runtime"]);
}

#[test]
fn test_object_can_implement_generic_interface() {
    let out = run_prints(r#"
        interface Transformer<T, U> {
            fun map(value: T): U
        }

        object IntToText : Transformer<Int, String> {
            override fun map(value: Int): String = "v" + value
        }

        fun emit(transformer: Transformer<Int, String>, value: Int): String {
            return transformer.map(value)
        }

        fun main() {
            println(emit(IntToText, 3))
        }
    "#);
    assert_eq!(out, &["v3"]);
}

#[test]
fn test_object_with_private_state_exposes_only_public_api() {
    let out = run_prints(r#"
        object Registry {
            private var next = 0
            fun next(): Int {
                next += 1
                return next
            }
        }

        fun main() {
            println(Registry.next())
            println(Registry.next())
        }
    "#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_nested_object_in_class_is_shared_across_instances() {
    let out = run_prints(r#"
        class Holder {
            object Cache {
                var value = 0
            }
        }

        fun main() {
            Holder.Cache.value += 1
            Holder.Cache.value += 2
            println(Holder.Cache.value)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_object_expression_as_interface_provider_is_type_stable() {
    let out = run_prints(r#"
        interface Printer {
            fun print(): String
        }

        fun makePrinter(prefix: String): Printer {
            return object : Printer {
                override fun print(): String = prefix + "!"
            }
        }

        fun main() {
            val first = makePrinter("a")
            val second = makePrinter("b")
            println(first.print())
            println(second.print())
        }
    "#);
    assert_eq!(out, &["a!", "b!"]);
}

#[test]
fn test_object_expression_interface_results_are_distinct_instances() {
    let out = run_prints(r#"
        interface Counter {
            fun next(): Int
        }

        fun makeCounter(start: Int): Counter {
            var value = start
            return object : Counter {
                override fun next(): Int {
                    value += 1
                    return value
                }
            }
        }

        fun main() {
            val first = makeCounter(0) as Any
            val second = makeCounter(0) as Any
            println((first as Counter).next())
            println((second as Counter).next())
            println(first === second)
        }
    "#);
    assert_eq!(out, &["1", "1", "false"]);
}

#[test]
fn test_object_expression_can_satisfy_multiple_interfaces_at_once() {
    let out = run_prints(r#"
        interface Named {
            fun name(): String
        }

        interface Valued {
            fun value(): Int
        }

        fun make(flag: String): Any {
            return object : Named, Valued {
                override fun name(): String = flag
                override fun value(): Int = 7
            }
        }

        fun main() {
            val item = make("ok")
            println((item as Named).name())
            println((item as Valued).value())
        }
    "#);
    assert_eq!(out, &["ok", "7"]);
}

#[test]
fn test_object_can_delegate_to_map_behavior() {
    let out = run_prints(r#"
        object Cache : Map<String, Int> by mapOf("a" to 1, "b" to 2) {
            val keysText = keys.joinToString("-")
        }

        fun main() {
            println(Cache["a"])
            println(Cache.keysText)
            println(Cache.size)
        }
    "#);
    assert_eq!(out, &["1", "a-b", "2"]);
}

#[test]
fn test_object_can_be_used_as_factory_for_function_type() {
    let out = run_prints(r#"
        object Builder : (String, String) -> String {
            override fun invoke(left: String, right: String): String {
                return left + right
            }
        }

        fun main() {
            val value: (String, String) -> String = Builder
            println(value("a", "b"))
            println(Builder.invoke("c", "d"))
        }
    "#);
    assert_eq!(out, &["ab", "cd"]);
}

#[test]
fn test_object_with_init_like_function_call_order_is_deterministic() {
    let out = run_prints(r#"
        object Log {
            val value: Int
            init {
                value = 10
            }

            fun value(): Int = value
        }

        fun main() {
            println(Log.value)
            println(Log.value())
        }
    "#);
    assert_eq!(out, &["10", "10"]);
}

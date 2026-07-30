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

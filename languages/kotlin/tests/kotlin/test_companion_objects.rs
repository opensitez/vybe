use crate::helpers::run_prints;

#[test]
fn test_companion_object_counter_tracks_instance_creations() {
    let out = run_prints(r#"
        class Token {
            companion object {
                var total = 0
            }

            init {
                Token.total += 1
            }
        }

        fun main() {
            Token()
            Token()
            Token()
            println(Token.total)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_companion_object_factory_returns_instances() {
    let out = run_prints(r#"
        class Widget private constructor(val label: String) {
            companion object {
                fun create(label: String): Widget = Widget(label)
            }
        }

        fun main() {
            val first = Widget.create("a")
            val second = Widget.create("b")
            println(first.label)
            println(second.label)
        }
    "#);
    assert_eq!(out, &["a", "b"]);
}

#[test]
fn test_companion_access_through_outer_name_is_stable() {
    let out = run_prints(r#"
        class Counter {
            companion object {
                val start = 5
            }
        }

        fun main() {
            println(Counter.start)
            println(Counter.Companion.start)
        }
    "#);
    assert_eq!(out, &["5", "5"]);
}

#[test]
fn test_companion_object_with_internal_state_and_mutation() {
    let out = run_prints(r#"
        class Store {
            companion object {
                private var next: Int = 0
                fun take(): Int {
                    next += 1
                    return next
                }
            }
        }

        fun main() {
            println(Store.take())
            println(Store.take())
        }
    "#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_companion_method_uses_its_own_properties() {
    let out = run_prints(r#"
        class Calculator {
            companion object {
                private const val scale = 10
                fun scaled(value: Int): Int = value * scale
            }
        }

        fun main() {
            println(Calculator.scaled(3))
        }
    "#);
    assert_eq!(out, &["30"]);
}

#[test]
fn test_companion_object_in_nested_class_is_addressable() {
    let out = run_prints(r#"
        class Holder {
            class Nested {
                companion object {
                    fun label(value: Int): String = "id:" + value
                }
            }
        }

        fun main() {
            println(Holder.Nested.label(7))
        }
    "#);
    assert_eq!(out, &["id:7"]);
}

#[test]
fn test_companion_object_with_init_block_runs_once() {
    let out = run_prints(r#"
        class Probe {
            companion object {
                var value = 0
                init {
                    value = 7
                }
            }
        }

        fun main() {
            println(Probe.value)
            println(Probe.value)
        }
    "#);
    assert_eq!(out, &["7", "7"]);
}

#[test]
fn test_companion_object_methods_can_return_receiver_instance() {
    let out = run_prints(r#"
        class Holder {
            val marker: String
            private constructor(marker: String) {
                this.marker = marker
            }

            companion object {
                fun create(): Holder = Holder("ok")
            }
        }

        fun main() {
            println(Holder.create().marker)
        }
    "#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_companion_object_shares_state_across_imported_instances() {
    let out = run_prints(r#"
        class Registry {
            companion object {
                var values = 0
            }
        }

        fun bump() {
            Registry.values += 1
        }

        fun main() {
            println(Registry.values)
            bump()
            bump()
            println(Registry.values)
        }
    "#);
    assert_eq!(out, &["0", "2"]);
}

#[test]
fn test_companion_with_extension_style_call_site() {
    let out = run_prints(r#"
        class Labeler {
            companion object {
                fun from(prefix: String, value: Int): String = prefix + value.toString()
            }
        }

        fun main() {
            println(Labeler.from("v", 4))
        }
    "#);
    assert_eq!(out, &["v4"]);
}

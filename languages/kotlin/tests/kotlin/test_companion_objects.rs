use crate::helpers::run_prints;

#[test]
fn test_companion_object_counter_tracks_instance_creations() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_companion_object_factory_returns_instances() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["a", "b"]);
}

#[test]
fn test_companion_access_through_outer_name_is_stable() {
    let out = run_prints(
        r#"
        class Counter {
            companion object {
                val start = 5
            }
        }

        fun main() {
            println(Counter.start)
            println(Counter.Companion.start)
        }
    "#,
    );
    assert_eq!(out, &["5", "5"]);
}

#[test]
fn test_companion_object_with_internal_state_and_mutation() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_companion_method_uses_its_own_properties() {
    let out = run_prints(
        r#"
        class Calculator {
            companion object {
                private const val scale = 10
                fun scaled(value: Int): Int = value * scale
            }
        }

        fun main() {
            println(Calculator.scaled(3))
        }
    "#,
    );
    assert_eq!(out, &["30"]);
}

#[test]
fn test_companion_object_in_nested_class_is_addressable() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["id:7"]);
}

#[test]
fn test_companion_object_companion_init_block_runs_once() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["7", "7"]);
}

#[test]
fn test_companion_object_methods_can_return_receiver_instance() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_companion_object_shares_state_across_imported_instances() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["0", "2"]);
}

#[test]
fn test_companion_with_extension_style_call_site() {
    let out = run_prints(
        r#"
        class Labeler {
            companion object {
                fun from(prefix: String, value: Int): String = prefix + value.toString()
            }
        }

        fun main() {
            println(Labeler.from("v", 4))
        }
    "#,
    );
    assert_eq!(out, &["v4"]);
}

#[test]
fn test_companion_object_isolated_state_per_host_type() {
    let out = run_prints(
        r#"
        class Left {
            companion object {
                var value = 1
            }
        }

        class Right {
            companion object {
                var value = 10
            }
        }

        fun main() {
            Left.value += 1
            Right.value += 5
            println(Left.value)
            println(Right.value)
        }
    "#,
    );
    assert_eq!(out, &["2", "15"]);
}

#[test]
fn test_companion_object_stores_cached_lookup_results() {
    let out = run_prints(
        r#"
        class Dictionary {
            companion object {
                private val cache = mutableMapOf<String, String>()

                fun put(key: String, value: String) {
                    cache[key] = value
                }

                fun lookup(key: String): String {
                    return cache[key] ?: ""
                }

                fun count(): Int = cache.size
            }
        }

        fun main() {
            println(Dictionary.count())
            Dictionary.put("x", "one")
            Dictionary.put("y", "two")
            println(Dictionary.lookup("x"))
            println(Dictionary.count())
        }
    "#,
    );
    assert_eq!(out, &["0", "one", "2"]);
}

#[test]
fn test_companion_object_can_implement_an_interface() {
    let out = run_prints(
        r#"
        interface Stamp {
            fun stamp(value: String): String
        }

        class Tagger {
            companion object : Stamp {
                override fun stamp(value: String): String = "tagged-" + value
            }
        }

        fun main() {
            println(Tagger.stamp("a"))
            println(Tagger.stamp("b"))
        }
    "#,
    );
    assert_eq!(out, &["tagged-a", "tagged-b"]);
}

#[test]
fn test_companion_object_uses_named_instance_reference() {
    let out = run_prints(
        r#"
        class Counter {
            companion object Factory {
                private var next = 0

                fun take(): Int {
                    next += 1
                    return next
                }
            }
        }

        fun main() {
            val first = Counter.Factory.take()
            val second = Counter.take()
            println(first)
            println(second)
        }
    "#,
    );
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_companion_object_with_init_block_runs_once() {
    let out = run_prints(
        r#"
        class Loader {
            companion object {
                var status = "cold"

                init {
                    status = "warm"
                }
            }
        }

        fun main() {
            println(Loader.status)
            println(Loader.status)
        }
    "#,
    );
    assert_eq!(out, &["warm", "warm"]);
}

#[test]
fn test_companion_object_exposes_computed_property() {
    let out = run_prints(
        r#"
        class Converter {
            companion object {
                const val base = 100
                val scaled: Int
                    get() = base * 2
            }
        }

        fun main() {
            println(Converter.base)
            println(Converter.scaled)
        }
    "#,
    );
    assert_eq!(out, &["100", "200"]);
}

#[test]
fn test_companion_object_factory_preserves_private_constructor_rules() {
    let out = run_prints(
        r#"
        class Token private constructor(val label: String) {
            companion object {
                fun create(prefix: String, suffix: Int): Token {
                    return Token(prefix + ":" + suffix.toString())
                }
            }
        }

        fun main() {
            println(Token.create("x", 9).label)
        }
    "#,
    );
    assert_eq!(out, &["x:9"]);
}

#[test]
fn test_companion_object_can_return_function_values() {
    let out = run_prints(
        r#"
        class Math {
            companion object {
                fun build(prefix: String): (Int) -> Int {
                    return { value -> value + prefix.length }
                }
            }
        }

        fun main() {
            val add = Math.build("hello")
            println(add(5))
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_generic_companion_factory_preserves_inferred_type() {
    let out = run_prints(
        r#"
        class Holder<T>(val value: T) {
            companion object {
                fun <T> make(value: T): Holder<T> = Holder(value)
            }
        }

        fun main() {
            val text = Holder.make("kotlin").value
            val number = Holder.make(12).value
            println(text)
            println(number)
        }
    "#,
    );
    assert_eq!(out, &["kotlin", "12"]);
}

#[test]
fn test_companion_calls_from_nested_types_share_parent_state() {
    let out = run_prints(
        r#"
        class Board {
            companion object {
                private var tokens = 0
                fun hit(): Int {
                    tokens += 1
                    return tokens
                }
            }

            class Checker {
                fun hit(): Int = Board.hit()
            }
        }

        fun main() {
            println(Board.hit())
            println(Board.Checker().hit())
            println(Board.hit())
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "3"]);
}

#[test]
fn test_named_companion_object_can_be_used_as_type() {
    let out = run_prints(
        r#"
        class Parser {
            companion object Validator {
                fun ok(value: String): Boolean = value.isNotEmpty()
            }
        }

        fun main() {
            val valid = Parser.Validator.ok("x")
            val invalid = Parser.Validator.ok("")
            println(valid)
            println(invalid)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_companion_object_accepts_top_level_helpers() {
    let out = run_prints(
        r#"
        fun stampPrefix(value: String): String = "[" + value + "]"

        class Packet {
            companion object {
                fun label(value: String): String = stampPrefix(value)
            }
        }

        fun main() {
            println(Packet.label("x"))
        }
    "#,
    );
    assert_eq!(out, &["[x]"]);
}

#[test]
fn test_companion_object_implements_function_type() {
    let out = run_prints(
        r#"
        class Prefixer {
            companion object : (String) -> String {
                override fun invoke(value: String): String = ">> " + value
            }
        }

        fun main() {
            val value: (String) -> String = Prefixer.Companion
            println(value("a"))
            println(Prefixer.Companion("b"))
        }
    "#,
    );
    assert_eq!(out, &[">> a", ">> b"]);
}

#[test]
fn test_companion_object_state_mutation_is_shared_with_factory_calls() {
    let out = run_prints(
        r#"
        class Counter {
            companion object {
                private var next = 0

                fun next(delta: Int = 1): Int {
                    next += delta
                    return next
                }

                fun current(): Int = next
            }
        }

        fun main() {
            println(Counter.next())
            println(Counter.next(3))
            println(Counter.next())
            println(Counter.current())
        }
    "#,
    );
    assert_eq!(out, &["1", "4", "5", "5"]);
}

#[test]
fn test_companion_object_init_only_runs_once_for_multiple_member_reads() {
    let out = run_prints(
        r#"
        var init_log = ""

        class Tracker {
            companion object {
                init {
                    init_log += "init;"
                }

                val tag = "ok"
            }
        }

        fun main() {
            println(init_log)
            println(Tracker.tag)
            println(Tracker.tag)
            println(init_log)
        }
    "#,
    );
    assert_eq!(out, &["", "ok", "ok", "init;"]);
}

#[test]
fn test_companion_object_can_be_used_as_an_interface_value() {
    let out = run_prints(
        r#"
        interface Named {
            fun name(): String
        }

        class Factory {
            companion object : Named {
                override fun name(): String = "factory"
            }
        }

        fun label(source: Named): String = source.name()

        fun main() {
            val source: Named = Factory.Companion
            println(label(source))
            println(label(Factory.Companion))
        }
    "#,
    );
    assert_eq!(out, &["factory", "factory"]);
}

#[test]
fn test_companion_object_can_implement_comparator_for_custom_sorting() {
    let out = run_prints(
        r#"
        data class Entry(val value: Int)

        class Holder {
            companion object : Comparator<Entry> {
                override fun compare(left: Entry, right: Entry): Int {
                    return right.value - left.value
                }
            }
        }

        fun main() {
            val values = listOf(Entry(1), Entry(3), Entry(2))
            val sorted = values.sortedWith(Holder.Companion)
            println(sorted.joinToString(",") { it.value.toString() })
        }
    "#,
    );
    assert_eq!(out, &["3,2,1"]);
}

#[test]
fn test_companion_object_for_nested_class_share_state_across_nested_instances() {
    let out = run_prints(
        r#"
        class Container {
            class Unit {
                companion object {
                    var count = 0
                    fun use(): Int {
                        count += 1
                        return count
                    }
                }
            }

            fun call(): Int = Unit.use()
        }

        fun main() {
            val one = Container.Unit.use()
            val two = Container.Unit()
            val three = two.call()
            val four = Container.Unit.use()
            println(one)
            println(three)
            println(four)
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "3"]);
}

#[test]
fn test_companion_object_can_store_private_initializer_output() {
    let out = run_prints(
        r#"
        class Builder {
            companion object {
                private const val prefix = "id:"
                val marker: String

                init {
                    marker = prefix + "1"
                }

                fun label(value: Int): String {
                    return marker + value.toString()
                }
            }
        }

        fun main() {
            println(Builder.label(4))
        }
    "#,
    );
    assert_eq!(out, &["id:14"]);
}

#[test]
fn test_companion_object_default_state_is_isolated_from_instance_state() {
    let out = run_prints(
        r#"
        class Holder {
            companion object {
                var global = 0
            }

            var local = 0

            init {
                local += 1
                global += local
            }
        }

        fun main() {
            val first = Holder()
            val second = Holder()
            println(first.local)
            println(second.local)
            println(Holder.global)
        }
    "#,
    );
    assert_eq!(out, &["1", "1", "2"]);
}

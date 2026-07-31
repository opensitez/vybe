use crate::helpers::run_prints;

#[test]
fn test_let_maps_non_null_value() {
    let out = run_prints(r#"
        fun main() {
            val value = 5
            val result = value.let { it + 7 }
            println(result)
        }
    "#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_let_with_explicit_parameter_name() {
    let out = run_prints(r#"
        fun main() {
            val source = "kotlin"
            val result = source.let { text -> text.uppercase() }
            println(result)
        }
    "#);
    assert_eq!(out, &["KOTLIN"]);
}

#[test]
fn test_let_chain_is_nested_transform() {
    let out = run_prints(r#"
        fun main() {
            val value = 3
                .let { it + 1 }
                .let { it * 10 }
            println(value)
        }
    "#);
    assert_eq!(out, &["40"]);
}

#[test]
fn test_let_on_nullable_returns_none_when_null() {
    let out = run_prints(r#"
        fun main() {
            val value: Int? = null
            val mapped = value?.let { it + 1 }
            println(mapped == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_nullable_let_with_fallback() {
    let out = run_prints(r#"
        fun main() {
            val value: Int? = null
            val mapped = value?.let { it + 1 } ?: -1
            println(mapped)
        }
    "#);
    assert_eq!(out, &["-1"]);
}

#[test]
fn test_let_to_scope_mutable_receiver_like_block() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            val doubled = values.let {
                val out = mutableListOf<Int>()
                for (value in it) {
                    out.add(value * 2)
                }
                out
            }
            println(doubled.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,4,6"]);
}

#[test]
fn test_run_block_executes_and_returns_value() {
    let out = run_prints(r#"
        fun main() {
            val total = run {
                val first = 4
                val second = 6
                first + second
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_run_block_can_capture_outer_variables() {
    let out = run_prints(r#"
        fun main() {
            var total = 1
            val value = run {
                total += 2
                total * 2
            }
            println(total)
            println(value)
        }
    "#);
    assert_eq!(out, &["3", "6"]);
}

#[test]
fn test_run_can_use_receiver_style_string_method_chain() {
    let out = run_prints(r#"
        fun main() {
            val result = "scoping".run {
                uppercase()
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["SCOPING"]);
}

#[test]
fn test_with_accesses_member_properties() {
    let out = run_prints(r#"
        class Box {
            val prefix = "left"
            val suffix = "right"
        }

        fun main() {
            val text = with(Box()) {
                prefix + ":" + suffix
            }
            println(text)
        }
    "#);
    assert_eq!(out, &["left:right"]);
}

#[test]
fn test_with_can_return_calculated_result() {
    let out = run_prints(r#"
        class Range(val start: Int, val end: Int)

        fun main() {
            val width = with(Range(2, 6)) {
                end - start
            }
            println(width)
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_apply_mutates_and_returns_receiver() {
    let out = run_prints(r#"
        fun main() {
            val text = StringBuilder("ko").apply {
                append("t")
                append("l")
                append("in")
            }
            println(text.toString())
            println(text.length)
        }
    "#);
    assert_eq!(out, &["kotlin", "6"]);
}

#[test]
fn test_apply_makes_inline_mutations_on_list() {
    let out = run_prints(r#"
        fun main() {
            val list = mutableListOf(1).apply {
                add(2)
                add(3)
            }
            println(list.joinToString("|"))
            println(list.size)
        }
    "#);
    assert_eq!(out, &["1|2|3", "3"]);
}

#[test]
fn test_apply_on_custom_type() {
    let out = run_prints(r#"
        class Counter {
            var value: Int = 0
            fun bump(step: Int) { value += step }
        }

        fun main() {
            val counter = Counter().apply {
                bump(2)
                bump(3)
            }
            println(counter.value)
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_also_records_side_effect_and_returns_receiver() {
    let out = run_prints(r#"
        fun main() {
            val events = mutableListOf<String>()
            val values = mutableListOf(1, 2).also {
                events.add("size-" + it.size.toString())
                it.add(3)
            }
            println(values.joinToString(","))
            println(events.joinToString("|"))
        }
    "#);
    assert_eq!(out, &["1,2,3", "size-2"]);
}

#[test]
fn test_also_chain_keeps_reference() {
    let out = run_prints(r#"
        fun main() {
            val log = mutableListOf<String>()
            val values = mutableListOf(10)
                .also { log.add("initial-" + it.size.toString()) }
                .also { it.add(20) }
                .also { log.add("after-" + it.size.toString()) }
            println(values.joinToString(";"))
            println(log.joinToString(","))
        }
    "#);
    assert_eq!(out, &["10;20", "initial-1,after-2"]);
}

#[test]
fn test_scoping_also_for_logging_without_mutation() {
    let out = run_prints(r#"
        fun main() {
            val base = 3
            val result = base.let { it * 2 }
                .also { }
            println(result)
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_take_if_keeps_matching_values() {
    let out = run_prints(r#"
        fun main() {
            println(10.takeIf { it > 5 })
            println(10.takeIf { it % 2 == 0 })
        }
    "#);
    assert_eq!(out, &["10", "10"]);
}

#[test]
fn test_take_if_returns_null_for_non_matching() {
    let out = run_prints(r#"
        fun main() {
            println(4.takeIf { it > 10 } == null)
            println(4.takeIf { it < 10 } == null)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_take_unless_rejects_matching_predicate() {
    let out = run_prints(r#"
        fun main() {
            println(7.takeUnless { it == 7 } == null)
            println(7.takeUnless { it > 10 } == null)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_take_unless_keeps_non_matching() {
    let out = run_prints(r#"
        fun main() {
            val value = 7.takeUnless { it == 0 }
            println(value)
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_scoped_chain_let_then_apply_then_also() {
    let out = run_prints(r#"
        fun main() {
            val log = mutableListOf<String>()
            val result = "ok".let { it.uppercase() }
                .also { log.add("a") }
                .let { it + "-done" }
                .also { log.add("b") }
            println(result)
            println(log.joinToString("-"))
        }
    "#);
    assert_eq!(out, &["OK-done", "a-b"]);
}

#[test]
fn test_scoped_chain_mix_with_take_if_for_filtering() {
    let out = run_prints(r#"
        fun main() {
            val base = "value".takeIf { it.length > 2 } ?: "none"
            val result = base.let { it + "-ok" }
            println(result)
        }
    "#);
    assert_eq!(out, &["value-ok"]);
}

#[test]
fn test_run_with_takeunless_for_default() {
    let out = run_prints(r#"
        fun main() {
            val candidate = "x".takeUnless { it.isEmpty() } ?: "missing"
            val result = candidate.run {
                "found:" + this
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["found:x"]);
}

#[test]
fn test_apply_with_nested_with_reads_its_properties() {
    let out = run_prints(r#"
        class Item {
            val id = 1
            val prefix = "item"
        }

        fun main() {
            val text = Item().apply {
                id.toString()
            }.let { it.id }
            println(text)
            val withText = with(Item()) { "$prefix-$id" }
            println(withText)
        }
    "#);
    assert_eq!(out, &["1", "item-1"]);
}

#[test]
fn test_scope_expression_with_multiple_receivers() {
    let out = run_prints(r#"
        class Holder {
            fun make(prefix: String): String = with(this) { "$prefix:value" }
        }

        fun main() {
            val value = Holder().run {
                make("start").also { }
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["start:value"]);
}

#[test]
fn test_apply_and_run_builder_style() {
    let out = run_prints(r#"
        class Builder {
            var text = ""
            fun build(): String = text
        }

        fun main() {
            val result = Builder()
                .apply { text = "a" }
                .apply { text += "b" }
                .run { build() }
            println(result)
        }
    "#);
    assert_eq!(out, &["ab"]);
}

#[test]
fn test_run_reads_and_transforms_receiver() {
    let out = run_prints(r#"
        class Holder {
            var value = 2
        }

        fun main() {
            val result = Holder().run {
                value *= 3
                value + 1
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_also_keeps_original_receiver_for_further_use() {
    let out = run_prints(r#"
        class Holder(var text: String)

        fun main() {
            val item = Holder("x")
                .also { it.text = it.text + "y" }
                .also { it.text = it.text + "z" }
            println(item.text)
        }
    "#);
    assert_eq!(out, &["xyz"]);
}

#[test]
fn test_with_block_scopes_multiple_property_updates() {
    let out = run_prints(r#"
        class Holder {
            var value = 1
            fun add(step: Int) { value += step }
        }

        fun main() {
            val out = with(Holder()) {
                add(3)
                add(2)
                value
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_let_returns_receiver_when_no_transformation() {
    let out = run_prints(r#"
        fun main() {
            val value = "kotlin"
            val result = value.let { it }
            println(result)
            println(result === value)
        }
    "#);
    assert_eq!(out, &["kotlin", "true"]);
}

#[test]
fn test_apply_keeps_reference_and_allows_multiple_mutations() {
    let out = run_prints(r#"
        class Box(var value: Int = 0)

        fun main() {
            val box = Box()
                .apply { value += 1 }
                .apply { value += 2 }
                .apply { value *= 2 }
            println(box.value)
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_run_can_catch_and_continue_from_exception_in_scope() {
    let out = run_prints(r#"
        fun main() {
            val out = try {
                run {
                    throw RuntimeException("boom")
                }
            } catch (error: RuntimeException) {
                "caught"
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_also_preserves_identity_with_side_effect_chain() {
    let out = run_prints(r#"
        fun main() {
            val first = mutableListOf(1)
            val second = first
                .also { it.add(2) }
                .also { it.add(3) }
            println(first === second)
            println(second.joinToString("|"))
        }
    "#);
    assert_eq!(out, &["true", "1|2|3"]);
}

#[test]
fn test_with_supports_reassigning_outer_mutable_state() {
    let out = run_prints(r#"
        class Holder(var value: Int)

        fun main() {
            var mutable = Holder(4)
            val label = with(mutable) {
                value *= 3
                "v" + value
            }
            println(label)
            println(mutable.value)
        }
    "#);
    assert_eq!(out, &["v12", "12"]);
}

#[test]
fn test_take_if_predicate_called_before_returning_value() {
    let out = run_prints(r#"
        fun main() {
            var checks = 0
            val value = 7.takeIf {
                checks++
                it == 7
            }
            println(value)
            println(checks)
        }
    "#);
    assert_eq!(out, &["7", "1"]);
}

#[test]
fn test_take_unless_predicate_not_called_for_null_receiver() {
    let out = run_prints(r#"
        fun main() {
            val value: Int? = null
            val result = value?.takeUnless {
                println("should-not-see-this")
                false
            }
            println(result == null)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_let_can_be_used_for_conditional_mapping() {
    let out = run_prints(r#"
        fun main() {
            val value = 8
            val result = if (value > 5) {
                value.let { it * 2 }
            } else {
                0
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["16"]);
}

#[test]
fn test_let_shadowed_name_does_not_escape_scope() {
    let out = run_prints(r#"
        fun main() {
            val value = "outer"
            val projected = value.let { value ->
                value.uppercase()
            }
            println(projected)
            println(value)
        }
    "#);
    assert_eq!(out, &["OUTER", "outer"]);
}

#[test]
fn test_let_with_nullable_receiver_preserves_null_short_circuit() {
    let out = run_prints(r#"
        fun main() {
            val value: String? = null
            val projected = value?.let {
                "inside"
            }
            println(projected == null)
            println(value)
        }
    "#);
    assert_eq!(out, &["true", "null"]);
}

#[test]
fn test_with_returns_its_block_value_not_context_object() {
    let out = run_prints(r#"
        class Holder(var count: Int)

        fun main() {
            val out = with(Holder(1)) {
                count += 9
                "count:" + count
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["count:10"]);
}

#[test]
fn test_run_keeps_outer_reference_unchanged_after_block() {
    let out = run_prints(r#"
        class State(var value: Int)

        fun main() {
            val state = State(4)
            val block = state.run {
                value = value * 3
                value + 1
            }
            println(state.value)
            println(block)
        }
    "#);
    assert_eq!(out, &["12", "13"]);
}

#[test]
fn test_apply_mutates_and_returns_receiver_with_multiple_property_writes() {
    let out = run_prints(r#"
        class Packet {
            var head: String = ""
            var tail: String = ""
        }

        fun main() {
            val packet = Packet().apply {
                head = "h"
                tail = "t"
                head += "+"
            }
            println(packet.head)
            println(packet.tail)
        }
    "#);
    assert_eq!(out, &["h+", "t"]);
}

#[test]
fn test_also_allows_side_effect_without_mutation() {
    let out = run_prints(r#"
        class Logger {
            val events = mutableListOf<String>()
        }

        fun main() {
            val value = Logger().also {
                it.events.add("created")
                it.events.add("ready")
            }
            println(value.events.joinToString(","))
        }
    "#);
    assert_eq!(out, &["created,ready"]);
}

#[test]
fn test_also_keeps_original_object_for_mutation_checks() {
    let out = run_prints(r#"
        class Holder(var total: Int)

        fun main() {
            val original = Holder(5)
            val observed = original.also {
                it.total += 10
            }
            println(original.total)
            println(original === observed)
        }
    "#);
    assert_eq!(out, &["15", "true"]);
}

#[test]
fn test_take_if_mutable_object_returns_same_reference() {
    let out = run_prints(r#"
        class Box(var n: Int)

        fun main() {
            val box = Box(3)
            val out = box.takeIf { it.n == 3 }
            println(out === box)
            println(out?.n)
        }
    "#);
    assert_eq!(out, &["true", "3"]);
}

#[test]
fn test_take_if_rejects_via_predicate_on_reference_state() {
    let out = run_prints(r#"
        class Box(var n: Int)

        fun main() {
            val value = Box(2)
            val filtered = value.takeIf { it.n > 2 }
            println(filtered == null)
            println(value.n)
        }
    "#);
    assert_eq!(out, &["true", "2"]);
}

#[test]
fn test_take_unless_on_reference_predicate() {
    let out = run_prints(r#"
        class Box(var n: Int)

        fun main() {
            val value = Box(11)
            val filtered = value.takeUnless { it.n % 2 == 1 }
            println(filtered == null)
            println(value.n)
            val keep = Box(4).takeUnless { it.n % 2 == 1 }
            println(keep?.n)
        }
    "#);
    assert_eq!(out, &["true", "11", "4"]);
}

#[test]
fn test_take_if_and_take_unless_chain_expresses_filtering_pipeline() {
    let out = run_prints(r#"
        class Box(var n: Int)

        fun main() {
            val value = Box(7)
            val result = value
                .takeIf { it.n > 5 }
                ?.takeUnless { it.n % 2 == 0 }
                ?.n ?: -1
            println(result)
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_let_can_collect_predicate_outcome_count() {
    let out = run_prints(r#"
        fun main() {
            var checks = 0
            val value = 3
            val doubled = value.let {
                checks++
                it * 2
            }
            println(doubled)
            println(checks)
        }
    "#);
    assert_eq!(out, &["6", "1"]);
}

#[test]
fn test_run_supports_returning_receiver_after_mutation_with_lambda() {
    let out = run_prints(r#"
        class State(var value: Int)

        fun main() {
            val source = State(1)
            val copy = source.run {
                value = value + 4
                this
            }
            println(source.value)
            println(copy.value)
            println(source === copy)
        }
    "#);
    assert_eq!(out, &["5", "5", "true"]);
}

#[test]
fn test_scoping_chain_handles_exception_and_recovers_via_try() {
    let out = run_prints(r#"
        fun main() {
            val result = try {
                5.let {
                    if (it < 0) throw RuntimeException("x")
                    it * 2
                }
            } catch (error: RuntimeException) {
                -1
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_nested_scoping_functions_show_receiver_visibility() {
    let out = run_prints(r#"
        class Counter(var value: Int)

        fun main() {
            val out = Counter(1).apply {
                this.value += 1
                val local = Counter(value).also {
                    it.value += 4
                }
                this.value += local.value
            }
            println(out.value)
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_with_on_nested_type_reads_both_rece_ivers() {
    let out = run_prints(r#"
        class Holder {
            var value = "h"
        }

        fun main() {
            val list = mutableListOf("x", "y")
            val result = with(list) {
                with(Holder()) {
                    list.add("z")
                    value = list.first()
                    value
                }
            }
            println(result)
            println(list.size)
        }
    "#);
    assert_eq!(out, &["x", "3"]);
}

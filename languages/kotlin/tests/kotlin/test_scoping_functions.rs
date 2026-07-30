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

use crate::helpers::run_prints;

#[test]
fn test_typealias_for_aliasing_simple_type() {
    let out = run_prints(r#"
        typealias Text = String

        fun main() {
            val value: Text = "ok"
            println(value)
        }
    "#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_typealias_for_class_constructor_arguments() {
    let out = run_prints(r#"
        data class PairValue(val id: Int, val label: String)
        typealias Entry = PairValue

        fun main() {
            val item: Entry = Entry(3, "x")
            println(item.id)
            println(item.label)
        }
    "#);
    assert_eq!(out, &["3", "x"]);
}

#[test]
fn test_typealias_for_function_type() {
    let out = run_prints(r#"
        typealias Reducer = (Int, Int) -> Int

        fun combine(value: Int, other: Int, op: Reducer): Int {
            return op(value, other)
        }

        fun main() {
            val result = combine(4, 5, { a, b -> a + b })
            println(result)
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_typealias_for_generic_type() {
    let out = run_prints(r#"
        typealias BoxOfInt = MutableList<Int>

        fun main() {
            val values: BoxOfInt = mutableListOf(1, 2, 3)
            values.add(4)
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_typealias_for_union_style_nesting() {
    let out = run_prints(r#"
        typealias Name = String
        typealias Named = Name

        fun main() {
            val value: Named = "x"
            println(value)
        }
    "#);
    assert_eq!(out, &["x"]);
}

#[test]
fn test_typealias_generic_holder_round_trip() {
    let out = run_prints(r#"
        class Holder<T>(val value: T)
        typealias StringHolder = Holder<String>

        fun main() {
            val item: StringHolder = StringHolder("done")
            println(item.value)
        }
    "#);
    assert_eq!(out, &["done"]);
}

#[test]
fn test_typealias_does_not_duplicate_runtime_identity() {
    let out = run_prints(r#"
        typealias Label = String

        class Box(var value: Label)

        fun main() {
            val first = Box("a")
            val second: Label = "a"
            first.value = second
            println(first.value)
        }
    "#);
    assert_eq!(out, &["a"]);
}

#[test]
fn test_typealias_for_map_type() {
    let out = run_prints(r#"
        typealias StringNumberMap = Map<String, Int>

        fun main() {
            val values: StringNumberMap = mapOf("a" to 1, "b" to 2)
            println(values["a"])
            println(values["c"] == null)
        }
    "#);
    assert_eq!(out, &["1", "true"]);
}

#[test]
fn test_local_typealias_scopes_to_block() {
    let out = run_prints(r#"
        fun make(): String {
            typealias LocalText = String
            val value: LocalText = "block"
            return value
        }

        fun main() {
            println(make())
        }
    "#);
    assert_eq!(out, &["block"]);
}

#[test]
fn test_typealias_with_array_type() {
    let out = run_prints(r#"
        typealias IntArrayLike = Array<Int>

        fun main() {
            val values: IntArrayLike = arrayOf(1, 2)
            println(values.size)
            println(values[0] + values[1])
        }
    "#);
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_typealias_for_tuple_like_type() {
    let out = run_prints(r#"
        typealias PairText = Pair<Int, String>

        fun main() {
            val value: PairText = Pair(4, "x")
            println(value.first)
            println(value.second)
        }
    "#);
    assert_eq!(out, &["4", "x"]);
}

#[test]
fn test_typealias_for_nullable_text_preserves_null_semantics() {
    let out = run_prints(r#"
        typealias Text = String?

        fun main() {
            val value: Text = null
            println(value == null)
            println((value ?: "fallback"))
        }
    "#);
    assert_eq!(out, &["true", "fallback"]);
}

#[test]
fn test_typealias_for_function_type_invocation() {
    let out = run_prints(r#"
        typealias Join = (String, String) -> String

        fun main() {
            val joiner: Join = { left, right -> left + right }
            println(joiner("a", "b"))
        }
    "#);
    assert_eq!(out, &["ab"]);
}

#[test]
fn test_typealias_for_generic_list_alias() {
    let out = run_prints(r#"
        typealias StringList = List<String>

        fun main() {
            val names: StringList = listOf("a", "b", "c")
            println(names.size)
            println(names[1])
        }
    "#);
    assert_eq!(out, &["3", "b"]);
}

#[test]
fn test_typealias_for_generic_container_reuses_type_parameter() {
    let out = run_prints(r#"
        typealias Container<T> = MutableList<T>

        fun main() {
            val values: Container<Int> = mutableListOf(1, 2)
            values.add(3)
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_typealias_for_mutable_map_key_and_value_types() {
    let out = run_prints(r#"
        typealias StringToIntMap = MutableMap<String, Int>

        fun main() {
            val values: StringToIntMap = mutableMapOf("a" to 1)
            values["b"] = 2
            println(values["a"])
            println(values["b"])
        }
    "#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_typealias_for_set_with_factory_function() {
    let out = run_prints(r#"
        typealias NameSet = HashSet<String>

        fun make(): NameSet {
            return NameSet(listOf("x", "y", "x"))
        }

        fun main() {
            println(make().size)
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_typealias_for_pair_projection() {
    let out = run_prints(r#"
        typealias PairLike = Pair<Int, Boolean>

        fun main() {
            val value: PairLike = Pair(4, true)
            println(value.first)
            println(value.second)
        }
    "#);
    assert_eq!(out, &["4", "true"]);
}

#[test]
fn test_typealias_for_interface_contract() {
    let out = run_prints(r#"
        interface Handler {
            fun run(value: Int): Int
        }

        typealias IncrHandler = Handler

        object Adder : IncrHandler {
            override fun run(value: Int): Int = value + 1
        }

        fun apply(handler: IncrHandler, value: Int): Int = handler.run(value)

        fun main() {
            println(apply(Adder, 8))
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_typealias_for_array_like_nested_type() {
    let out = run_prints(r#"
        typealias Matrix = Array<IntArray>

        fun main() {
            val grid: Matrix = arrayOf(intArrayOf(1, 2), intArrayOf(3, 4))
            println(grid.size)
            println(grid[1][0])
        }
    "#);
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_local_typealias_can_rebind_inner_type() {
    let out = run_prints(r#"
        fun main() {
            typealias LocalMap = Map<String, Int>
            val values: LocalMap = mapOf("x" to 10)
            println(values["x"])
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_typealias_can_alias_result_of_function_type() {
    let out = run_prints(r#"
        typealias Next = () -> Int

        fun make(value: Int): Next {
            return { value + 1 }
        }

        fun main() {
            val next: Next = make(7)
            println(next())
        }
    "#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_typealias_for_receiver_function_type() {
    let out = run_prints(r#"
        typealias Formatter = String.() -> String

        val shout: Formatter = { this.uppercase() + "!" }

        fun main() {
            println("k".shout())
        }
    "#);
    assert_eq!(out, &["K!"]);
}

#[test]
fn test_typealias_for_multi_arg_mapper() {
    let out = run_prints(r#"
        typealias Joiner = (String, String, String) -> String

        fun join(parts: Joiner): String {
            return parts("a", "b", "c")
        }

        fun main() {
            val joiner: Joiner = { left, middle, right -> left + "-" + middle + "-" + right }
            println(join(joiner))
        }
    "#);
    assert_eq!(out, &["a-b-c"]);
}

#[test]
fn test_typealias_for_generic_factory_function() {
    let out = run_prints(r#"
        data class Box<T>(val value: T)

        typealias BoxFactory<T> = (T) -> Box<T>

        fun make(value: Int, factory: BoxFactory<Int>): Box<Int> {
            return factory(value)
        }

        fun main() {
            val boxFactory: BoxFactory<Int> = { Box(it * 2) }
            println(make(4, boxFactory).value)
        }
    "#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_typealias_for_mutable_set_specializes_numeric_key() {
    let out = run_prints(r#"
        typealias NumberSet = MutableSet<Int>

        fun main() {
            val values: NumberSet = hashSetOf(1, 2, 2, 3)
            values.add(3)
            println(values.size)
            println(values.contains(2))
        }
    "#);
    assert_eq!(out, &["3", "true"]);
}

#[test]
fn test_typealias_for_map_with_list_payloads() {
    let out = run_prints(r#"
        typealias ScoresByLabel = MutableMap<String, MutableList<Int>>

        fun main() {
            val scores: ScoresByLabel = mutableMapOf()
            scores["a"] = mutableListOf(1, 2)
            scores["a"]?.add(3)
            println(scores["a"]?.size)
            println(scores["a"]?.sum())
        }
    "#);
    assert_eq!(out, &["3", "6"]);
}

#[test]
fn test_typealias_for_comparator_reference() {
    let out = run_prints(r#"
        typealias IntSort = Comparator<Int>

        fun main() {
            val values = mutableListOf(4, 1, 3, 2)
            val order: IntSort = Comparator { left, right -> right - left }
            println(values.sortedWith(order).joinToString("-"))
        }
    "#);
    assert_eq!(out, &["4-3-2-1"]);
}

#[test]
fn test_typealias_for_pair_projection_shape() {
    let out = run_prints(r#"
        typealias PairAlias = Pair<String, Int>

        fun main() {
            val value: PairAlias = Pair("x", 7)
            println(value.first)
            println(value.second)
        }
    "#);
    assert_eq!(out, &["x", "7"]);
}

#[test]
fn test_typealias_for_nullable_function_type() {
    let out = run_prints(r#"
        typealias OptionalText = (() -> String)?

        fun main() {
            val value: OptionalText = null
            println((value == null))
            println((value?.invoke() ?: "empty"))
        }
    "#);
    assert_eq!(out, &["true", "empty"]);
}

#[test]
fn test_typealias_nested_aliases_remain_assignable() {
    let out = run_prints(r#"
        typealias BaseLabel = String
        typealias UserLabel = BaseLabel
        typealias DisplayLabel = UserLabel

        fun main() {
            val source: BaseLabel = "admin"
            val alias: DisplayLabel = source
            val roundTrip: UserLabel = alias
            println(alias)
            println(roundTrip)
        }
    "#);
    assert_eq!(out, &["admin", "admin"]);
}

#[test]
fn test_typealias_generic_alias_for_pair_collections() {
    let out = run_prints(r#"
        typealias PairList<T> = List<Pair<T, T>>

        fun total(values: PairList<Int>): Int {
            return values.fold(0) { acc, item -> acc + item.first + item.second }
        }

        fun main() {
            val values: PairList<Int> = listOf(Pair(1, 2), Pair(3, 4))
            println(total(values))
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_typealias_receiver_function_invocation_from_aliased_builder() {
    let out = run_prints(r#"
        typealias StringBuilderRecipe = StringBuilder.() -> Unit

        fun formatValue(value: String): String {
            val recipe: StringBuilderRecipe = {
                append(value)
                append("-done")
            }
            val target = StringBuilder()
            target.recipe()
            return target.toString()
        }

        fun main() {
            println(formatValue("ok"))
        }
    "#);
    assert_eq!(out, &["ok-done"]);
}

#[test]
fn test_typealias_extension_function_on_aliased_set_type() {
    let out = run_prints(r#"
        typealias NameSet = MutableSet<String>

        fun NameSet.sortedSignature(): String {
            return this.toList().sorted().joinToString("|")
        }

        fun main() {
            val names: NameSet = hashSetOf("z", "a", "m")
            names.add("c")
            println(names.sortedSignature())
        }
    "#);
    assert_eq!(out, &["a|c|m|z"]);
}

#[test]
fn test_typealias_sequence_operations_preserve_lazy_contract() {
    let out = run_prints(r#"
        typealias IntSequence = Sequence<Int>

        fun firstSquares(limit: Int): IntSequence {
            return generateSequence(0) { value ->
                if (value + 2 <= limit) value + 2 else null
            }
        }

        fun main() {
            val values = firstSquares(6)
            println(values.take(3).joinToString(","))
            println(firstSquares(6).sum())
        }
    "#);
    assert_eq!(out, &["2,4,6", "12"]);
}

#[test]
fn test_typealias_map_entry_is_a_type_projection_only() {
    let out = run_prints(r#"
        typealias ScoreEntry = Map.Entry<String, Int>

        fun main() {
            val map = mapOf("a" to 10, "b" to 20)
            val top: ScoreEntry = map.entries.reduce { acc, item ->
                if (item.value > acc.value) item else acc
            }
            println(top.key)
            println(top.value)
        }
    "#);
    assert_eq!(out, &["b", "20"]);
}

#[test]
fn test_typealias_for_generic_bounded_functions() {
    let out = run_prints(r#"
        typealias ComparableList<T> = List<T>

        fun <T : Comparable<T>> maxOfList(values: ComparableList<T>): T {
            return values.maxOrNull()!!
        }

        fun main() {
            val words: ComparableList<String> = listOf("bb", "aaa", "c")
            println(maxOfList(words))
        }
    "#);
    assert_eq!(out, &["c"]);
}

#[test]
fn test_typealias_for_java_collection_type() {
    let out = run_prints(r#"
        typealias JavaMap = java.util.LinkedHashMap<String, Int>

        fun main() {
            val counts: JavaMap = java.util.LinkedHashMap<String, Int>()
            counts["a"] = 1
            counts["b"] = 2
            counts.put("a", 3)
            println(counts["a"])
            println(counts.size)
        }
    "#);
    assert_eq!(out, &["3", "2"]);
}

#[test]
fn test_typealias_nested_function_type_pipeline() {
    let out = run_prints(r#"
        typealias Transformer<T> = (T) -> T
        typealias IntTransformer = Transformer<Int>

        fun applyTwice(value: Int, first: IntTransformer, second: IntTransformer): Int {
            return second(first(value))
        }

        fun main() {
            val stepA: IntTransformer = { it + 5 }
            val stepB: IntTransformer = { it * 2 }
            println(applyTwice(3, stepA, stepB))
        }
    "#);
    assert_eq!(out, &["16"]);
}

#[test]
fn test_typealias_aliases_with_nullable_generic_type() {
    let out = run_prints(r#"
        typealias Maybe<T> = T?
        typealias MaybeText = Maybe<String>

        fun fallback(value: MaybeText, default: String): String {
            return value ?: default
        }

        fun main() {
            val present: MaybeText = "ok"
            val missing: MaybeText = null
            println(fallback(present, "none"))
            println(fallback(missing, "none"))
        }
    "#);
    assert_eq!(out, &["ok", "none"]);
}

#[test]
fn test_typealias_local_in_generic_function_context() {
    let out = run_prints(r#"
        typealias Wrapped<T> = List<T>

        fun <T> describe(values: Wrapped<T>): String {
            typealias FirstLabel = String
            val first: FirstLabel = values.firstOrNull().toString()
            return first
        }

        fun main() {
            val names = listOf("kotlin", "tests")
            println(describe(names))
            val numbers = listOf(1, 2, 3)
            println(describe(numbers))
        }
    "#);
    assert_eq!(out, &["kotlin", "1"]);
}

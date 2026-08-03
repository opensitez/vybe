kotlin_run_cases! {
    test_associate_pairs_to_map => (r##"
        fun main() {
            val map = listOf("a" to 1, "b" to 2, "c" to 3).associateBy { it.first }
            println(map.size)
            println(map["b"])
        }
    "##, &[
        "3",
        "2",
    ]),
    test_associate_with_values => (r##"
        fun main() {
            val words = listOf("cat", "dog", "eel")
            val map = words.associateWith { it.length }
            println(map["cat"])
            println(map["dog"])
        }
    "##, &[
        "3",
        "3",
    ]),
    test_associate_by_length => (r##"
        fun main() {
            val words = listOf("aa", "bb", "c", "ddd")
            val map = words.associateBy { it.length }
            println(map[1])
            println(map[3])
        }
    "##, &[
        "c",
        "ddd",
    ]),
    test_associate_by_collision_last_wins => (r##"
        fun main() {
            val items = listOf("k1" to 1, "k1" to 2, "k2" to 3)
            val map = items.associateBy { it.first }
            println(map.size)
            println(map["k1"])
        }
    "##, &[
        "2",
        "2",
    ]),
    test_associate_by_transform_value => (r##"
        fun main() {
            val list = listOf("x", "yy", "zzz")
            val map = list.associateBy({ it }, { it.length })
            println(map["x"])
            println(map["zzz"])
        }
    "##, &[
        "1",
        "3",
    ]),
    test_associate_with_to_mutable => (r##"
        fun main() {
            val out = mutableMapOf<String, Int>()
            listOf("a", "bb", "ccc").associateWithTo(out) { it.length }
            println(out["bb"])
            println(out.size)
        }
    "##, &[
        "2",
        "3",
    ]),
    test_associate_by_to_mutable => (r##"
        fun main() {
            val out = mutableMapOf<Int, String>()
            listOf("a", "bb", "ccc").associateByTo(out, { it.length }, { it })
            println(out[1])
            println(out[2])
            println(out[3])
        }
    "##, &[
        "a",
        "bb",
        "ccc",
    ]),
    test_associate_by_to_overwrites_duplicates => (r##"
        fun main() {
            val out = linkedHashMapOf<Int, String>()
            listOf("ab", "ac", "b").associateByTo(out, { it.length }, { it })
            println(out.size)
            println(out[2])
        }
    "##, &[
        "2",
        "b",
    ]),
    test_associate_with_to_with_filter => (r##"
        fun main() {
            val out = mutableMapOf<String, Int>()
            listOf("x", "yy", "yyy", "yyyy").filter { it.length > 1 }.associateWithTo(out) { it.length }
            println(out.size)
            println(out["yy"])
            println(out.containsKey("x"))
        }
    "##, &[
        "3",
        "2",
        "false",
    ]),
    test_associate_pairs_from_array => (r##"
        fun main() {
            val entries = arrayOf(Pair("a", 1), Pair("b", 2), Pair("c", 3))
            val map = entries.toMap()
            println(map.keys.joinToString(","))
            println(map.values.sum())
        }
    "##, &[
        "a,b,c",
        "6",
    ]),
    test_associate_with_index => (r##"
        fun main() {
            val values = listOf("a", "b", "c")
            val map = values.associateBy({ it }, { it.toInt() - 96 })
            println(map["a"])
            println(map["c"])
        }
    "##, &[
        "97",
        "99",
    ]),
    test_associate_with_duplicate_keys_in_to_map => (r##"
        fun main() {
            val values = listOf(Pair("k", 1), Pair("k", 2), Pair("k", 3))
            val map = values.toMap()
            println(map.size)
            println(map["k"])
        }
    "##, &[
        "1",
        "3",
    ]),
    test_associate_by_string_codes => (r##"
        fun main() {
            val map = listOf("a", "b", "c").associateBy({ it.first() }, { it.code })
            println(map.keys.joinToString(","))
            println(map["a"])
        }
    "##, &[
        "a,b,c",
        "97",
    ]),
    test_associate_with_unicode => (r##"
        fun main() {
            val map = listOf("aa", "ß", "é").associateWith { it.length }
            println(map["ß"])
            println(map["é"])
        }
    "##, &[
        "1",
        "1",
    ]),
    test_associate_by_sorted => (r##"
        fun main() {
            val map = listOf("dog", "ant", "cat", "zoo").associateBy { it.length }
            val keys = map.keys.toList().sorted()
            println(keys.joinToString(","))
        }
    "##, &[
        "3,4",
    ]),
    test_associate_with_from_empty => (r##"
        fun main() {
            val map = listOf<String>().associateWith { it.length }
            println(map.isEmpty())
            println(map.size)
        }
    "##, &[
        "true",
        "0",
    ]),
    test_associate_by_with_mutable_output_reuse => (r##"
        fun main() {
            val out = mutableMapOf<Int, String>()
            listOf("x", "yy", "zzz").associateByTo(out, { it.length }) { it }
            listOf("a", "bb").associateByTo(out, { it.length }) { it + "!" }
            println(out[1])
            println(out[2])
        }
    "##, &[
        "x",
        "bb!",
    ]),
    test_associate_projection_pairs => (r##"
        fun main() {
            val map = mapOf(1 to "one", 2 to "two")
            val keys = map.keys.associateWith { it * 10 }
            println(keys[1])
            println(keys[2])
        }
    "##, &[
        "10",
        "20",
    ]),
    test_associate_with_count_chars => (r##"
        fun main() {
            val input = listOf("a", "bc", "de", "f")
            val map = input.associateWith { it.count() }
            val total = map.values.sum()
            println(map.size)
            println(total)
        }
    "##, &[
        "4",
        "7",
    ]),
    test_associate_by_boolean_key => (r##"
        fun main() {
            val map = listOf(1, 2, 3, 4).associateBy { it % 2 == 0 }
            println(map.keys.joinToString(","))
            println(map[true]?.size)
        }
    "##, &[
        "false,true",
        "2",
    ]),
    test_associate_with_even_markers => (r##"
        fun main() {
            val map = listOf(1, 2, 3, 4).associateWith { if (it % 2 == 0) "E" else "O" }
            println(map[1])
            println(map[2])
            println(map[4])
        }
    "##, &[
        "O",
        "E",
        "E",
    ]),
    test_associate_from_sequence => (r##"
        fun main() {
            val map = sequenceOf("a", "bb", "ccc").associateBy { it.length }
            println(map[1])
            println(map[3])
        }
    "##, &[
        "a",
        "ccc",
    ]),
    test_associate_to_mutable_seeded => (r##"
        fun main() {
            val out = linkedMapOf<Int, String>()
            out[1] = "seed"
            listOf("alpha", "bee").associateByTo(out, { it.length }) { it }
            println(out[5])
            println(out.size)
        }
    "##, &[
        "bee",
        "2",
    ]),
    test_associate_with_from_chars => (r##"
        fun main() {
            val map = listOf('a', 'b', 'c').associateWith { it.code }
            println(map['a'])
            println(map['c'])
        }
    "##, &[
        "97",
        "99",
    ]),
    test_associate_projection_chain => (r##"
        fun main() {
            val map = listOf("aa", "bb", "ccc").associateBy { it[0] }.mapValues { it.value.length }
            println(map.keys.joinToString(","))
            println(map["a"])
            println(map["c"])
        }
    "##, &[
        "a,b,c",
        "2",
        "3",
    ]),
    test_associate_with_index_then_to_map => (r##"
        fun main() {
            val map = (0..3).associateWith { it * 2 }
            println(map[2])
            println(map.values.sum())
        }
    "##, &[
        "4",
        "12",
    ]),
    test_associate_by_filter_by_value => (r##"
        fun main() {
            val map = listOf("x", "yy", "zzz").associateBy { it.length }.filterValues { it.startsWith("z") }
            println(map.size)
            println(map[3])
        }
    "##, &[
        "1",
        "zzz",
    ]),
    test_associate_by_to_existing_map_grows => (r##"
        fun main() {
            val out = mutableMapOf<Int, String>()
            out[99] = "seed"
            listOf("a", "bb").associateByTo(out, { it.length }, { it })
            println(out[99])
            println(out[1])
            println(out[2])
        }
    "##, &[
        "seed",
        "a",
        "bb",
    ]),
    test_associate_with_to_sorted_result => (r##"
        fun main() {
            val out = mutableMapOf<Int, String>()
            listOf("x", "yy", "zzz").associateWithTo(out) { it.length.toString() }
            val keys = out.keys.toList().sorted()
            println(keys.joinToString(","))
            println(out[2])
        }
    "##, &[
        "1,2,3",
        "2",
    ]),
    test_associate_empty_map_behavior => (r##"
        fun main() {
            val map = emptyList<Int>().associateBy { it }
            println(map.isEmpty())
            println(map.size)
        }
    "##, &[
        "true",
        "0",
    ]),
    test_associate_with_prefix => (r##"
        fun main() {
            val map = listOf(1, 2, 3).associateWith { "id-" + it }
            println(map[1])
            println(map[2])
            println(map[3])
        }
    "##, &[
        "id-1",
        "id-2",
        "id-3",
    ]),
    test_associate_by_to_joined_keys => (r##"
        fun main() {
            val out = linkedMapOf<Int, String>()
            listOf("one", "two", "ten").associateByTo(out, { it.length }, { it.first() })
            println(out.keys.joinToString(","))
            println(out[3])
        }
    "##, &[
        "3,",
        "t",
    ]) }

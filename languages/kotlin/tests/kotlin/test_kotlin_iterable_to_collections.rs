kotlin_run_cases! {
    test_to_list_from_array => (r##"
        fun main() {
            val arr = arrayOf(1, 2, 3)
            val out = arr.asList()
            println(out.joinToString(","))
        }
    "##, vec![String::from("1,2,3")]),
    test_to_mutable_list => (r##"
        fun main() {
            val out = listOf(1, 2, 3).toMutableList()
            out.add(4)
            println(out.joinToString(","))
        }
    "##, vec![String::from("1,2,3,4")]),
    test_to_set_behavior => (r##"
        fun main() {
            val out = listOf(1, 1, 2, 2, 3).toSet()
            println(out.size)
            println(out.joinToString(","))
        }
    "##, vec![String::from("3"), String::from("1,2,3")]),
    test_to_sorted_set => (r##"
        fun main() {
            val out = listOf(3, 1, 2).toSortedSet()
            println(out.joinToString(","))
        }
    "##, vec![String::from("1,2,3")]),
    test_to_set_from_sequence => (r##"
        fun main() {
            val out = sequenceOf(1, 2, 2, 3).toSet()
            println(out.size)
            println(out.joinToString(","))
        }
    "##, vec![String::from("3"), String::from("1,2,3")]),
    test_to_map_from_pairs => (r##"
        fun main() {
            val pairs = listOf("a" to 1, "b" to 2)
            val out = pairs.toMap()
            println(out.size)
            println(out["a"].toString())
            println(out["b"].toString())
        }
    "##, vec![String::from("2"), String::from("1"), String::from("2")]),
    test_map_not_null_to_map => (r##"
        fun main() {
            val values = listOf("a" to 1, "b" to 2)
            val keys = values.map { it.first }
            val valuesOut = values.map { it.second }
            println(keys.joinToString(","))
            println(valuesOut.joinToString(","))
        }
    "##, vec![String::from("a,b"), String::from("1,2")]),
    test_associate_with_index => (r##"
        fun main() {
            val values = listOf("x", "y", "z")
            val out = values.associateWith { it.length + 1 }
            println(out.size)
            println(out["y"].toString())
        }
    "##, vec![String::from("3"), String::from("2")]),
    test_associate_by => (r##"
        fun main() {
            val values = listOf("aa", "b", "ccc")
            val out = values.associateBy { it.length }
            println(out[1])
            println(out[2])
            println(out[3])
        }
    "##, vec![String::from("b"), String::from("aa"), String::from("ccc")]),
    test_associate_by_to => (r##"
        fun main() {
            val values = listOf("aa", "bb", "c")
            val map = linkedMapOf<Int, String>()
            values.associateByTo(map) { it.length }
            println(map.keys.joinToString(","))
            println(map[2])
        }
    "##, vec![String::from("2,1"), String::from("bb")]),
    test_to_typed_array => (r##"
        fun main() {
            val ints = listOf(1, 2, 3).toIntArray()
            println(ints.joinToString(","))
            val chars = listOf('a', 'b').toCharArray()
            println(chars.joinToString(","))
        }
    "##, vec![String::from("1,2,3"), String::from("a,b")]),
    test_as_reversed_list => (r##"
        fun main() {
            val m = listOf(1, 2, 3).asReversed()
            println(m.joinToString(","))
        }
    "##, vec![String::from("3,2,1")]),
    test_iterable_to_hash_set => (r##"
        fun main() {
            val s = listOf(1, 2, 1, 3).toHashSet()
            println(s.contains(2).toString())
            println(s.size)
        }
    "##, vec![String::from("true"), String::from("3")]),
    test_iterable_zip_to_set => (r##"
        fun main() {
            val zipped = (1..3).zip("abc")
            val s = zipped.toSet()
            println(s.size)
            println(s.any { it.first == 2 }.toString())
        }
    "##, vec![String::from("3"), String::from("true")]),
    test_iterable_to_sequence => (r##"
        fun main() {
            val seq = listOf(1, 2, 3).asSequence()
            println(seq.sum().toString())
        }
    "##, vec![String::from("6")]),
    test_iterable_map_to_set => (r##"
        fun main() {
            val words = listOf("ab", "c", "de")
            val lengths = words.map { it.length }.toSet()
            println(lengths.joinToString(","))
        }
    "##, vec![String::from("2,1")]),
    test_iterable_flatten_map => (r##"
        fun main() {
            val outer = listOf(listOf(1, 2), listOf(2, 3), listOf(3))
            val flat = outer.flatten()
            val unique = flat.toSet()
            println(flat.joinToString(","))
            println(unique.joinToString(","))
        }
    "##, vec![String::from("1,2,2,3,3"), String::from("1,2,3")]) }

kotlin_run_cases! {
    test_boolean_array_sizes => (r#"
        fun main() {
            val data = booleanArrayOf(true, false, true)
            println(data.size)
        }
    "#, vec![String::from("3")]),
    test_boolean_array_indexing => (r#"
        fun main() {
            val data = booleanArrayOf(true, false, true)
            println(data[0].toString())
            println(data[2].toString())
        }
    "#, vec![String::from("true"), String::from("true")]),
    test_boolean_array_mutation => (r#"
        fun main() {
            val data = booleanArrayOf(true, false)
            data[1] = true
            println(data[1].toString())
        }
    "#, vec![String::from("true")]),
    test_boolean_array_loop_count_true => (r#"
        fun main() {
            val data = booleanArrayOf(true, false, true, true)
            var count = 0
            for (v in data) {
                if (v) { count = count + 1 }
            }
            println(count)
        }
    "#, vec![String::from("3")]),
    test_boolean_array_negate_loop => (r#"
        fun main() {
            val data = booleanArrayOf(true, false)
            var bits = ""
            for (v in data) {
                bits = bits + if (v) "1" else "0"
            }
            println(bits)
        }
    "#, vec![String::from("10")]),
    test_boolean_array_to_list_map => (r#"
        fun main() {
            val data = booleanArrayOf(true, true, false)
            var ones = ""
            for (i in data.indices) {
                ones = ones + data[i].toString()
                if (i + 1 < data.size) { ones = ones + "," }
            }
            println(ones)
        }
    "#, vec![String::from("true,true,false")]),
    test_boolean_array_for_each => (r#"
        fun main() {
            val data = booleanArrayOf(false, true)
            var total = ""
            data.forEach { item ->
                total = total + item.toString()
            }
            println(total)
        }
    "#, vec![String::from("falsetrue")]),
    test_boolean_array_copy => (r#"
        fun main() {
            val source = booleanArrayOf(true, false)
            val copy = source.copyOf()
            copy[0] = false
            println(source[0].toString())
            println(copy[0].toString())
        }
    "#, vec![String::from("true"), String::from("false")]),
    test_boolean_array_equals_reference => (r#"
        fun main() {
            val a = booleanArrayOf(true, false)
            val b = booleanArrayOf(true, false)
            println((a == b).toString())
        }
    "#, vec![String::from("false")]),
    test_boolean_array_empty => (r#"
        fun main() {
            val empty = booleanArrayOf()
            println(empty.size)
            println(empty.isEmpty().toString())
        }
    "#, vec![String::from("0"), String::from("true")]),
    test_boolean_array_mixed => (r#"
        fun main() {
            val data = booleanArrayOf(true, false, false, true)
            var i = 0
            var out = ""
            while (i < data.size) {
                out = out + if (data[i]) "T" else "F"
                i = i + 1
            }
            println(out)
        }
    "#, vec![String::from("TFFT")]),
    test_boolean_array_nested_loop => (r#"
        fun main() {
            val a = booleanArrayOf(true, false)
            var out = ""
            for (x in a) {
                for (y in a) {
                    out = out + if (x && y) "1" else "0"
                }
            }
            println(out)
        }
    "#, vec![String::from("1000")]),
    // ^ x && y in pair order (t,t)(t,f)(f,t)(f,f) → 1,0,0,0 (real Kotlin
    // agrees).
}

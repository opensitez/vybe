kotlin_run_cases! {
    test_short_array_sizes => (r#"
        fun main() {
            val data = shortArrayOf(10, 20)
            println(data.size)
        }
    "#, vec![String::from("2")]),
    test_short_array_access => (r#"
        fun main() {
            val data = shortArrayOf(10, 20)
            println(data[1].toString())
        }
    "#, vec![String::from("20")]),
    test_short_array_mutation => (r#"
        fun main() {
            val data = shortArrayOf(1, 2)
            data[1] = 9
            println(data[1].toString())
        }
    "#, vec![String::from("9")]),
    test_short_array_sum => (r#"
        fun main() {
            val data = shortArrayOf(4, 5, 6)
            var total: Int = 0
            for (x in data) { total = total + x.toInt() }
            println(total)
        }
    "#, vec![String::from("15")]),
    test_short_array_range_loop => (r#"
        fun main() {
            val data = shortArrayOf(1, 2, 3)
            var out = ""
            for (i in data.indices) {
                out = out + data[i].toString()
            }
            println(out)
        }
    "#, vec![String::from("123")]),
    test_short_array_copy => (r#"
        fun main() {
            val a = shortArrayOf(1, 2)
            val b = a.copyOf()
            b[0] = 10
            println(a[0].toString())
            println(b[0].toString())
        }
    "#, vec![String::from("1"), String::from("10")]),
    test_short_array_empty => (r#"
        fun main() {
            val empty = shortArrayOf()
            println(empty.size)
            println(empty.isEmpty().toString())
        }
    "#, vec![String::from("0"), String::from("true")]),
    test_short_array_min => (r#"
        fun main() {
            val a = shortArrayOf(9, -3, 7)
            var min = a[0].toInt()
            for (i in 1 until a.size) {
                if (a[i].toInt() < min) { min = a[i].toInt() }
            }
            println(min)
        }
    "#, vec![String::from("-3")]),
    test_short_array_contains => (r#"
        fun main() {
            val a = shortArrayOf(1, 2, 3)
            var found = false
            for (x in a) { if (x.toInt() == 2) { found = true } }
            println(found.toString())
        }
    "#, vec![String::from("true")]),
    test_short_array_clone_ref => (r#"
        fun main() {
            val a = shortArrayOf(2, 4, 6)
            val b = a
            b[1] = 9
            println(a[1].toString())
        }
    "#, vec![String::from("9")]),
    test_short_array_filter_positive => (r#"
        fun main() {
            val a = shortArrayOf(-1, 0, 2)
            var count = 0
            for (x in a) { if (x.toInt() > 0) { count = count + 1 } }
            println(count)
        }
    "#, vec![String::from("1")]),
    test_short_array_join_char => (r#"
        fun main() {
            val a = shortArrayOf(1, 2, 3)
            var out = ""
            for (i in a.indices) {
                out = out + a[i].toString()
                if (i + 1 < a.size) { out = out + "|" }
            }
            println(out)
        }
    "#, vec![String::from("1|2|3")]),
}

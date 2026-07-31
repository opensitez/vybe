kotlin_run_cases! {
    test_char_array_sizes => (r#"
        fun main() {
            val data = charArrayOf('a', 'b', 'c')
            println(data.size)
        }
    "#, vec![String::from("3")]),
    test_char_array_index => (r#"
        fun main() {
            val data = charArrayOf('a', 'b', 'c')
            println(data[0].toString())
            println(data[2].toString())
        }
    "#, vec![String::from("a"), String::from("c")]),
    test_char_array_update => (r#"
        fun main() {
            val data = charArrayOf('a', 'b')
            data[1] = 'z'
            println(data[1].toString())
        }
    "#, vec![String::from("z")]),
    test_char_array_concat => (r#"
        fun main() {
            val data = charArrayOf('x', 'y')
            var s = ""
            for (c in data) { s = s + c.toString() }
            println(s)
        }
    "#, vec![String::from("xy")]),
    test_char_array_reverse => (r#"
        fun main() {
            val data = charArrayOf('a', 'b', 'c')
            var out = ""
            for (i in data.indices.reversed()) {
                out = out + data[i].toString()
            }
            println(out)
        }
    "#, vec![String::from("cba")]),
    test_char_array_contains => (r#"
        fun main() {
            val data = charArrayOf('a', 'b', 'c')
            var hit = false
            for (c in data) { if (c == 'b') { hit = true } }
            println(hit.toString())
        }
    "#, vec![String::from("true")]),
    test_char_array_copy => (r#"
        fun main() {
            val a = charArrayOf('a', 'b')
            val b = a.copyOf()
            b[0] = 'x'
            println(a[0].toString())
            println(b[0].toString())
        }
    "#, vec![String::from("a"), String::from("x")]),
    test_char_array_empty => (r#"
        fun main() {
            val e = charArrayOf()
            println(e.size)
            println(e.isEmpty().toString())
        }
    "#, vec![String::from("0"), String::from("true")]),
    test_char_array_indexed_loop => (r#"
        fun main() {
            val a = charArrayOf('a', 'b', 'c')
            var out = ""
            var i = 0
            while (i < a.size) {
                out = out + a[i].toString()
                i = i + 1
            }
            println(out)
        }
    "#, vec![String::from("abc")]),
    test_char_array_pair => (r#"
        fun main() {
            val a = charArrayOf('a', 'b')
            val b = charArrayOf('a', 'b')
            println((a == b).toString())
        }
    "#, vec![String::from("false")]),
    test_char_array_upper_hint => (r#"
        fun main() {
            val a = charArrayOf('a', 'b', 'c')
            var u = ""
            for (c in a) { u = u + c.toString().uppercase() }
            println(u)
        }
    "#, vec![String::from("ABC")]),
    test_char_array_find_first => (r#"
        fun main() {
            val a = charArrayOf('x', 'y', 'z')
            var pos = -1
            for (i in a.indices) {
                if (a[i] == 'y') { pos = i }
            }
            println(pos)
        }
    "#, vec![String::from("1")]),
}

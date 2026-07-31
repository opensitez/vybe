kotlin_run_cases! {
    test_byte_array_sizes => (r#"
        fun main() {
            val data = byteArrayOf(1, 2, 3)
            println(data.size)
        }
    "#, vec![String::from("3")]),
    test_byte_array_indexing => (r#"
        fun main() {
            val data = byteArrayOf(5, 7, 9)
            println(data[1].toString())
        }
    "#, vec![String::from("7")]),
    test_byte_array_set => (r#"
        fun main() {
            val data = byteArrayOf(1, 2)
            data[0] = 9
            println(data[0].toString())
        }
    "#, vec![String::from("9")]),
    test_byte_array_sum => (r#"
        fun main() {
            val data = byteArrayOf(1, 2, 3)
            var total: Int = 0
            for (v in data) { total = total + v.toInt() }
            println(total)
        }
    "#, vec![String::from("6")]),
    test_byte_array_loop_text => (r#"
        fun main() {
            val data = byteArrayOf(4, 5)
            var s = ""
            for (v in data) { s = s + v.toString() }
            println(s)
        }
    "#, vec![String::from("45")]),
    test_byte_array_copy_of => (r#"
        fun main() {
            val a = byteArrayOf(1, 2)
            val b = a.copyOf()
            b[0] = 8
            println(a[0].toString())
            println(b[0].toString())
        }
    "#, vec![String::from("1"), String::from("8")]),
    test_byte_array_fill => (r#"
        fun main() {
            val a = byteArrayOf(1, 1, 1)
            java.util.Arrays.fill(a, 3)
            println(a[0].toString())
            println(a[2].toString())
        }
    "#, vec![String::from("3"), String::from("3")]),
    test_byte_array_empty => (r#"
        fun main() {
            val empty = byteArrayOf()
            println(empty.size)
            println(empty.isEmpty().toString())
        }
    "#, vec![String::from("0"), String::from("true")]),
    test_byte_array_from_list => (r#"
        fun main() {
            val list = listOf<Byte>(1, 2, 3)
            val a = ByteArray(list.size) { list[it] }
            var out = ""
            for (x in a) { out = out + x.toString() }
            println(out)
        }
    "#, vec![String::from("123")]),
    test_byte_array_max => (r#"
        fun main() {
            val data = byteArrayOf(1, 5, 3)
            var max = data[0].toInt()
            for (i in 1 until data.size) {
                if (data[i].toInt() > max) { max = data[i].toInt() }
            }
            println(max)
        }
    "#, vec![String::from("5")]),
    test_byte_array_equals => (r#"
        fun main() {
            val a = byteArrayOf(1, 2)
            val b = byteArrayOf(1, 2)
            println((a == b).toString())
        }
    "#, vec![String::from("false")]),
    test_byte_array_find_index => (r#"
        fun main() {
            val a = byteArrayOf(6, 7, 8)
            var idx = -1
            for (i in a.indices) {
                if (a[i] == 7.toByte()) { idx = i }
            }
            println(idx)
        }
    "#, vec![String::from("1")]),
}

kotlin_run_cases! {
    test_int_progression_sum => (r##"
        fun main() {
            val values = 1..5
            var total = 0
            for (v in values) { total += v }
            println(total)
            println(values.step)
            println(values.last)
        }
    "##, vec![String::from("15"), String::from("1"), String::from("5")]),
    test_int_progression_down_to => (r##"
        fun main() {
            var out = ""
            for (v in 5 downTo 1) {
                out = out + v.toString()
            }
            println(out)
        }
    "##, vec![String::from("54321")]),
    test_int_progression_step => (r##"
        fun main() {
            val values = 1..10 step 2
            var out = ""
            for (v in values) {
                out = out + v.toString()
                out = out + ","
            }
            println(out)
        }
    "##, vec![String::from("1,3,5,7,9,")]),
    test_int_progression_step_down_to => (r##"
        fun main() {
            val values = 10 downTo 1 step 3
            var out = ""
            for (v in values) {
                out = out + v.toString()
                if (v > 1) out = out + ","
            }
            println(out)
        }
    "##, vec![String::from("10,7,4,1")]),
    test_int_progression_until => (r##"
        fun main() {
            val values = 1 until 4
            var out = ""
            for (v in values) { out = out + v.toString() }
            println(out)
            println(values.last)
        }
    "##, vec![String::from("123"), String::from("3")]),
    test_long_progression => (r##"
        fun main() {
            val values = 1L..5L
            println(values.start)
            println(values.endInclusive)
            println(values.step)
        }
    "##, vec![String::from("1"), String::from("5"), String::from("1")]),
    test_char_progression => (r##"
        fun main() {
            var out = ""
            for (c in 'a'..'d') { out = out + c }
            println(out)
            println(('c' in 'a'..'d').toString())
            println(('x' in 'a'..'d').toString())
        }
    "##, vec![String::from("abcd"), String::from("true"), String::from("false")]),
    test_contains_in_range => (r##"
        fun main() {
            println((5 in 1..10).toString())
            println((11 in 1..10).toString())
            println((1L in 1L..10L).toString())
            println((10L in 1L until 10L).toString())
        }
    "##, vec![String::from("true"), String::from("false"), String::from("true"), String::from("false")]),
    test_range_reversal => (r##"
        fun main() {
            val r = (1..5).reversed()
            println(r.first())
            println(r.last())
        }
    "##, vec![String::from("5"), String::from("1")]),
    test_progression_first_last => (r##"
        fun main() {
            val r = 10 downTo 2 step 4
            println(r.first)
            println(r.last)
            println(r.step)
        }
    "##, vec![String::from("10"), String::from("2"), String::from("-4")]),
    test_range_to_list => (r##"
        fun main() {
            val r = (1..3).toList()
            val a = (5 downTo 3).toList()
            println(r.joinToString(","))
            println(a.joinToString(","))
        }
    "##, vec![String::from("1,2,3"), String::from("5,4,3")]),
    test_progression_take_drop => (r##"
        fun main() {
            val r = (1..10)
            val first = r.take(4)
            val remain = r.drop(4)
            println(first.joinToString(","))
            println(remain.take(3).joinToString(","))
        }
    "##, vec![String::from("[1, 2, 3, 4]"), String::from("[5, 6, 7]")]),
    test_step_without_change => (r##"
        fun main() {
            val values = 1..10 step 1
            var x = 0
            for (v in values) { x = v }
            println(x)
            val empty = 10 downTo 12 step 2
            println(empty.toList().size)
        }
    "##, vec![String::from("10"), String::from("0")]),
    test_range_with_negative_start => (r##"
        fun main() {
            var out = ""
            for (v in -1..2) {
                out = out + v.toString()
            }
            println(out)
        }
    "##, vec![String::from("-1012")]),
    test_range_while_loop_equiv => (r##"
        fun main() {
            var i = 3
            val r = 1..3
            var out = ""
            while (i >= 1) {
                out = out + i.toString()
                i -= 1
            }
            println(r.toList().joinToString(","))
            println(out)
        }
    "##, vec![String::from("1,2,3"), String::from("321")]),
    test_char_range_reversed => (r##"
        fun main() {
            val out = ('c' downTo 'a').toList()
            println(out.toList().joinToString(","))
        }
    "##, vec![String::from("c,b,a")]),
}

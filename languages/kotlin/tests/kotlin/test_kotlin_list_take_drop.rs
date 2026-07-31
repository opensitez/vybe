kotlin_run_cases! {
    test_take_zero => (r#"
        fun main() {
            val a = listOf("a", "b", "c")
            println(a.take(0).size)
        }
    "#, vec![String::from("0")]),
    test_take_three => (r#"
        fun main() {
            val a = listOf("a", "b", "c", "d")
            println(a.take(3).toString())
        }
    "#, vec![String::from("[a, b, c]")]),
    test_take_more => (r#"
        fun main() {
            val a = listOf("a", "b")
            println(a.take(5).toString())
        }
    "#, vec![String::from("[a, b]")]),
    test_drop_zero => (r#"
        fun main() {
            val a = listOf("a", "b", "c")
            println(a.drop(0).toString())
        }
    "#, vec![String::from("[a, b, c]")]),
    test_drop_one => (r#"
        fun main() {
            val a = listOf("a", "b", "c")
            println(a.drop(1).toString())
        }
    "#, vec![String::from("[b, c]")]),
    test_drop_more => (r#"
        fun main() {
            val a = listOf("a", "b")
            println(a.drop(5).toString())
        }
    "#, vec![String::from("[]")]),
    test_take_last_zero => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.takeLast(0).toString())
        }
    "#, vec![String::from("[]")]),
    test_take_last_all => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.takeLast(5).toString())
        }
    "#, vec![String::from("[1, 2, 3]")]),
    test_drop_last_zero => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.dropLast(0).toString())
        }
    "#, vec![String::from("[1, 2, 3]")]),
    test_drop_last_two => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.dropLast(2).toString())
        }
    "#, vec![String::from("[1]")]),
    test_drop_last_more => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.dropLast(5).toString())
        }
    "#, vec![String::from("[]")]),
    test_take_drop_combo => (r#"
        fun main() {
            val a = listOf("a", "b", "c", "d")
            println(a.drop(1).take(2).toString())
        }
    "#, vec![String::from("[b, c]")]),
    test_take_last_drop_last => (r#"
        fun main() {
            val a = listOf("a", "b", "c", "d")
            println(a.take(3).dropLast(1).toString())
        }
    "#, vec![String::from("[a, b]")]),
}

kotlin_run_cases! {
    test_slice_prefix => (r#"
        fun main() {
            val a = listOf(1, 2, 3, 4)
            println(a.slice(0..1).toString())
        }
    "#, vec![String::from("[1, 2]")]),
    test_slice_range => (r#"
        fun main() {
            val a = listOf(1, 2, 3, 4)
            println(a.slice(1 until 3).toString())
        }
    "#, vec![String::from("[2, 3]")]),
    test_slice_tail => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.subList(1, a.size).toString())
        }
    "#, vec![String::from("[2, 3]")]),
    test_slice_out_of_bounds_safe => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.subList(0, a.size).toString())
        }
    "#, vec![String::from("[1, 2, 3]")]),
    test_list_take => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.take(2).toString())
        }
    "#, vec![String::from("[1, 2]")]),
    test_list_drop => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.drop(1).toString())
        }
    "#, vec![String::from("[2, 3]")]),
    test_list_take_last => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.takeLast(2).toString())
        }
    "#, vec![String::from("[2, 3]")]),
    test_list_drop_last => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.dropLast(1).toString())
        }
    "#, vec![String::from("[1, 2]")]),
    test_list_drop_last_zero => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.dropLast(0).toString())
        }
    "#, vec![String::from("[1, 2, 3]")]),
    test_list_take_empty => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            println(a.take(0).toString())
        }
    "#, vec![String::from("[]")]),
    test_list_slice_step => (r#"
        fun main() {
            val a = listOf(1, 2, 3, 4, 5)
            println(a.slice(0..4 step 2).toString())
        }
    "#, vec![String::from("[1, 3, 5]")]),
    test_list_slice_invalid => (r#"
        fun main() {
            val a = listOf(1, 2, 3)
            val empty = a.slice(0..-1)
            println(empty.toString())
        }
    "#, vec![String::from("[]")]) }

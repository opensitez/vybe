kotlin_run_cases! {
    test_do_while_minimum_runs => (r#"
        fun main() {
            var i = 5
            var out = 0
            do {
                out = out + i
                i = i - 1
            } while (i > 5)
            println(out)
        }
    "#, vec!["5"]),
    test_do_while_loop_count => (r#"
        fun main() {
            var i = 0
            var out = ""
            do {
                out = out + i.toString()
                i = i + 1
            } while (i < 3)
            println(out)
        }
    "#, vec!["012"]),
    test_do_while_with_continue => (r#"
        fun main() {
            var i = 0
            var out = ""
            do {
                i = i + 1
                if (i == 2) {
                    continue
                }
                out = out + i.toString()
            } while (i < 4)
            println(out)
        }
    "#, vec!["134"]),
}

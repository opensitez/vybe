kotlin_run_cases! {
    test_while_value_accum => (r#"
        fun main() {
            var i = 0
            var out = 0
            while (i < 3) {
                out = out + i
                i = i + 1
            }
            println(out)
        }
    "#, vec!["3"]),
    test_while_break => (r#"
        fun main() {
            var i = 0
            while (true) {
                if (i == 2) {
                    break
                }
                i = i + 1
            }
            println(i)
        }
    "#, vec!["2"]),
    test_while_nested_expression => (r#"
        fun main() {
            var i = 0
            var txt = ""
            while (i < 2) {
                txt = txt + i.toString()
                i = i + 1
            }
            println(txt)
        }
    "#, vec!["01"]),
}

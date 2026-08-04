kotlin_run_cases! {
    test_nested_labelled_break => (r#"
        fun main() {
            outer@ for (i in 1..3) {
                for (j in 1..3) {
                    if (i == 2 && j == 2) {
                        break@outer
                    }
                    println(i.toString() + "-" + j.toString())
                }
            }
        }
    "#, vec!["1-1", "1-2", "1-3", "2-1",]),
    test_nested_labelled_continue => (r#"
        fun main() {
            outer@ for (i in 1..2) {
                for (j in 1..3) {
                    if (j == 2) {
                        continue@outer
                    }
                    println(i.toString() + "-" + j.toString())
                }
            }
        }
    "#, vec!["1-1", "2-1", "2-3"]),
    test_nested_labeled_while => (r#"
        fun main() {
            var i = 0
            outer@ while (i < 3) {
                var j = 0
                while (j < 2) {
                    if (i == 1 && j == 0) {
                        j = j + 1
                        i = i + 1
                        continue@outer
                    }
                    println(i.toString() + ":" + j.toString())
                    j = j + 1
                }
                i = i + 1
            }
        }
    "#, vec!["0:0", "0:1", "2:0", "2:1"]),
}

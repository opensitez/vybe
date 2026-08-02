// vybe-test: kotlin/kotlin_while_results/test_while_nested_expression
// origin: languages/kotlin/tests/kotlin/test_kotlin_while_results.rs

fun main() {
            var i = 0
            var txt = ""
            while (i < 2) {
                txt = txt + i.toString()
                i = i + 1
            }
            println(txt)
        }


// vybe-test: kotlin/when_expressions/test_when_with_throws_in_subject_branches
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun classify(n: Int): String {
            return when (n) {
                0 -> "zero"
                1 -> "one"
                in 2..9 -> "few"
                else -> throw Error("too-large")
            }
        }

        fun main() {
            try {
                println(classify(1))
                println(classify(5))
                println(classify(20))
            } catch (e: Error) {
                println("error")
            }
        }


// vybe-test: kotlin/tuples/test_destructuring_works_with_triple_in_while_like_rewrite
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun main() {
            var index = 0
            var total = ""
            val values = listOf(Triple("a", 1, 10), Triple("b", 2, 20))
            while (index < values.size) {
                val (_, left, right) = values[index]
                total += "$left:$right;"
                index++
            }
            println(total)
        }


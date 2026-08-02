// vybe-test: kotlin/loops/test_for_range_descending_iterates_in_reverse
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var output = ""
            for (i in 5 downTo 2) {
                output += i.toString()
            }
            println(output)
        }


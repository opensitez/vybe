// vybe-test: kotlin/loops/test_for_over_list_of_strings
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            val words = listOf("a", "b", "c")
            println(words.joinToString(","))
            var joined = ""
            for (word in words) {
                joined += word
            }
            println(joined)
        }


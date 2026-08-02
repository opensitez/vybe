// vybe-test: kotlin/collections/test_array_of_booleans_and_count_true
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val flags = arrayOf(true, false, true, true)
            var trueCount = 0
            for (flag in flags) {
                if (flag) trueCount += 1
            }
            println(trueCount)
        }


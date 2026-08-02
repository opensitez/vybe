// vybe-test: kotlin/collection_fold_scan/test_running_reduce_sequence
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun main() {
            val out = listOf(2, 3, 4).runningReduce { acc, n -> acc * n }
            println(out.joinToString(","))
            val empty = emptyList<Int>()
            try {
                println(empty.runningReduce { a, b -> a + b }.joinToString(","))
            } catch (e: Exception) {
                println("err")
            }
        }


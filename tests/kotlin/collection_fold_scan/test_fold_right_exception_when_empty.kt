// vybe-test: kotlin/collection_fold_scan/test_fold_right_exception_when_empty
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun main() {
            val empty = emptyList<Int>()
            try {
                println(empty.reduceRight { a, b -> a + b })
            } catch (e: Exception) {
                println("no")
            }
        }


// vybe-test: kotlin/loops/test_while_condition_recomputed_each_iteration
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var i = 0
            var threshold = 3
            var total = 0
            while (i < threshold) {
                total += 1
                threshold += if (i == 1) 2 else 0
                i += 1
            }
            println(i)
            println(total)
        }


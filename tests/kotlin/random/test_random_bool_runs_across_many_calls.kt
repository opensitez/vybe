// vybe-test: kotlin/random/test_random_bool_runs_across_many_calls
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun main() {
            val r1 = kotlin.random.Random(107)
            val r2 = kotlin.random.Random(107)
            var trueCount = 0
            var i = 0
            while (i < 8) {
                if (r1.nextBoolean()) trueCount++
                i++
            }
            var trueCount2 = 0
            var j = 0
            while (j < 8) {
                if (r2.nextBoolean()) trueCount2++
                j++
            }
            println(trueCount == trueCount2)
            println(trueCount in 0..8)
        }


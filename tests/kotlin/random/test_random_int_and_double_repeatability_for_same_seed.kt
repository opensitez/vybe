// vybe-test: kotlin/random/test_random_int_and_double_repeatability_for_same_seed
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = kotlin.random.Random(113)
            val b = kotlin.random.Random(113)
            val aFirst = a.nextInt(1000)
            val aSecond = a.nextInt(1000)
            val bFirst = b.nextInt(1000)
            val bSecond = b.nextInt(1000)
            val aDouble = a.nextDouble()
            val bDouble = b.nextDouble()
            __check((aFirst == bFirst).toString(), "true")
            __check((aSecond == bSecond).toString(), "true")
            __check((aDouble == bDouble).toString(), "true")
        }

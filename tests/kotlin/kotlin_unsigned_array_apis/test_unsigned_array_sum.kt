// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun main() {
            val u = uintArrayOf(1u, 2u, 3u)
            val b = ubyteArrayOf(10u, 20u)
            var uSum = 0u
            var bSum = 0u
            for (x in u) { uSum += x }
            for (x in b) { bSum += x.toUInt() }
            println(uSum.toString())
            println(bSum.toString())
        }


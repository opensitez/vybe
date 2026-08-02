// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_all
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun main() {
            val u = uintArrayOf(1u, 4u, 9u)
            val b = ubyteArrayOf(1u, 2u)
            var allPositive = true
            for (x in u) { if (x == 0u) allPositive = false }
            var hasLow = false
            for (x in b) { if (x < 2u) hasLow = true }
            println(allPositive.toString())
            println(hasLow.toString())
        }


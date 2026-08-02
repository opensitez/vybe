// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_contains
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun main() {
            val u = uintArrayOf(7u, 8u, 9u)
            val b = ubyteArrayOf(1u, 2u, 3u)
            var found = false
            var missing = false
            for (x in u) { if (x == 8u) found = true }
            for (x in b) { if (x == 9u) missing = true }
            println(found.toString())
            println(missing.toString())
        }


// vybe-test: kotlin/kotlin_short_array_apis/test_short_array_contains
// origin: languages/kotlin/tests/kotlin/test_kotlin_short_array_apis.rs

fun main() {
            val a = shortArrayOf(1, 2, 3)
            var found = false
            for (x in a) { if (x.toInt() == 2) { found = true } }
            println(found.toString())
        }


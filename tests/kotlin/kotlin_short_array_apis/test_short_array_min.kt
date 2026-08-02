// vybe-test: kotlin/kotlin_short_array_apis/test_short_array_min
// origin: languages/kotlin/tests/kotlin/test_kotlin_short_array_apis.rs

fun main() {
            val a = shortArrayOf(9, -3, 7)
            var min = a[0].toInt()
            for (i in 1 until a.size) {
                if (a[i].toInt() < min) { min = a[i].toInt() }
            }
            println(min)
        }


// vybe-test: kotlin/kotlin_short_array_apis/test_short_array_join_char
// origin: languages/kotlin/tests/kotlin/test_kotlin_short_array_apis.rs

fun main() {
            val a = shortArrayOf(1, 2, 3)
            var out = ""
            for (i in a.indices) {
                out = out + a[i].toString()
                if (i + 1 < a.size) { out = out + "|" }
            }
            println(out)
        }


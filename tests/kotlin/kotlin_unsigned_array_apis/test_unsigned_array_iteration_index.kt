// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_iteration_index
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun main() {
            val s = ushortArrayOf(3u, 4u, 5u)
            var out = ""
            for (i in s.indices) {
                out = out + s[i].toString()
                if (i + 1 < s.size) { out = out + "," }
            }
            println(out)
        }


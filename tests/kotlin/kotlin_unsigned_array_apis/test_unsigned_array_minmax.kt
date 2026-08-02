// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_minmax
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun main() {
            val u = uintArrayOf(3u, 1u, 4u)
            val b = ubyteArrayOf(9u, 2u, 7u)
            var minU = u[0]
            var maxB = b[0].toInt()
            var i = 1
            while (i < u.size) {
                if (u[i] < minU) { minU = u[i] }
                i += 1
            }
            i = 0
            while (i < b.size) {
                val v = b[i].toInt()
                if (v > maxB) { maxB = v }
                i += 1
            }
            println(minU.toString())
            println(maxB.toString())
        }


// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_conversion_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun main() {
            val raw = uintArrayOf(1u, 2u)
            var i = 0
            var total = 0
            while (i < raw.size) {
                total += raw[i].toInt()
                i += 1
            }
            println(total)
            println(UByteArray(3) { 1u }.size)
            println(ULongArray(3) { 2uL }.joinToString(","))
        }


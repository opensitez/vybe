// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_builder
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = UIntArray(4) { it.toUInt() + 1u }
            val b = UByteArray(3) { (it + 10).toUByte() }
            val s = UShortArray(2) { ((it * 2 + 1).toUShort()) }
            val l = ULongArray(2) { (it.toULong() + 1uL) * 100uL }
            __check((u.joinToString(",")).toString(), "1,2,3,4")
            __check((b.joinToString(",")).toString(), "10,11,12")
            __check((s.joinToString(",")).toString(), "1,3")
            __check((l.joinToString(",")).toString(), "100,200")
        }

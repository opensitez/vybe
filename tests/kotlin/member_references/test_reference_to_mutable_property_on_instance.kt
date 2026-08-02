// vybe-test: kotlin/member_references/test_reference_to_mutable_property_on_instance
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Holder {
            var value: Int = 0
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val holder = Holder()
            holder.value = 8
            val read = holder::value
            holder.value = 1
            val read2 = Holder::value
            __check((read()).toString(), "1")
            __check((read2(holder)).toString(), "1")
        }

// vybe-test: kotlin/inheritance_dispatch/test_generic_dispatch_on_inheritance
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface Box {
            fun value(): Int
        }

        open class Holder<T : Box> : Box {
            override fun value(): Int = 0
        }

        class Fast : Holder<IntBox>() {
            override fun value(): Int = 4
        }

        class IntBox : Box {
            override fun value(): Int = 9
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Holder<*> = Fast()
            val typed: Box = Fast()
            __check((item.value()).toString(), "4")
            __check((typed.value()).toString(), "4")
        }

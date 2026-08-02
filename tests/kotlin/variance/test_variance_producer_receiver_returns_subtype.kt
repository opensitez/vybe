// vybe-test: kotlin/variance/test_variance_producer_receiver_returns_subtype
// origin: languages/kotlin/tests/kotlin/test_variance.rs

interface Produce<out T> { fun next(): T }
        class SpecialProducer : Produce<Special> {
            override fun next(): Special = Special()
        }
        open class Special
        open class Item
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p: Produce<Item> = SpecialProducer()
            __check((p.next() is Special).toString(), "true")
        }

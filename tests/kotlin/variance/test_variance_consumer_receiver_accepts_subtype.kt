// vybe-test: kotlin/variance/test_variance_consumer_receiver_accepts_subtype
// origin: languages/kotlin/tests/kotlin/test_variance.rs

interface Consume<in T> { fun accept(v: T) }
        open class Item
        class Special : Item()
        class Sink : Consume<Item> { override fun accept(v: Item) { __check((v::class.simpleName).toString(), "Special") } }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val consume: Consume<Special> = Sink()
            consume.accept(Special())
        }

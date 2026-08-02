// vybe-test: kotlin/iterator_protocol/test_custom_iterator_implementing_interface
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

class RangeIterator : Iterator<Int> {
            private var i = 0
            private val end = 3
            override fun hasNext() = i < end
            override fun next(): Int {
                val value = i
                i += 1
                return value
            }
        }

        class RangeIterable : Iterable<Int> {
            override fun iterator(): Iterator<Int> = RangeIterator()
        }

        fun main() {
            val it = RangeIterable().iterator()
            var sum = 0
            for (value in RangeIterable()) {
                sum += value
            }
            println(sum)
            println(it.hasNext())
            println(it.next())
            println(it.next())
        }


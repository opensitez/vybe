kotlin_run_cases! {
    test_unsigned_array_sizes => (r##"
        fun main() {
            val u = uintArrayOf(1u, 2u, 3u)
            val b = ubyteArrayOf(4u, 5u, 6u)
            val s = ushortArrayOf(7u, 8u, 9u)
            val l = ulongArrayOf(10uL, 11uL, 12uL)
            println(u.size)
            println(b.size)
            println(s.size)
            println(l.size)
        }
    "##, vec![String::from("3"), String::from("3"), String::from("3"), String::from("3")]),
    test_unsigned_array_indexing => (r##"
        fun main() {
            val u = uintArrayOf(10u, 20u, 30u)
            val b = ubyteArrayOf(11u, 12u)
            val s = ushortArrayOf(101u, 102u)
            val l = ulongArrayOf(1000uL, 2000uL)
            println(u[1].toString())
            println(b[0].toString())
            println(s[1].toString())
            println(l[1].toString())
        }
    "##, vec![String::from("20"), String::from("11"), String::from("102"), String::from("2000")]),
    test_unsigned_array_mutation => (r##"
        fun main() {
            val u = uintArrayOf(1u, 2u)
            val b = ubyteArrayOf(1u, 2u)
            u[0] = 9u
            b[1] = 8u
            println(u[0].toString())
            println(b[1].toString())
        }
    "##, vec![String::from("9"), String::from("8")]),
    test_unsigned_array_builder => (r##"
        fun main() {
            val u = UIntArray(4) { it.toUInt() + 1u }
            val b = UByteArray(3) { (it + 10).toUByte() }
            val s = UShortArray(2) { ((it * 2 + 1).toUShort()) }
            val l = ULongArray(2) { (it.toULong() + 1uL) * 100uL }
            println(u.joinToString(","))
            println(b.joinToString(","))
            println(s.joinToString(","))
            println(l.joinToString(","))
        }
    "##, vec![String::from("1,2,3,4"), String::from("10,11,12"), String::from("1,3"), String::from("100,200")]),
    test_unsigned_array_copy_of => (r##"
        fun main() {
            val u = uintArrayOf(2u, 4u, 6u)
            val copy = u.copyOf()
            copy[0] = 99u
            println(u[0].toString())
            println(copy[0].toString())
        }
    "##, vec![String::from("2"), String::from("99")]),
    test_unsigned_array_empty => (r##"
        fun main() {
            val u = uintArrayOf()
            println(u.size)
            println(u.isEmpty().toString())
            println(ubyteArrayOf().size)
            println(ushortArrayOf().size)
            println(ulongArrayOf().size)
        }
    "##, vec![String::from("0"), String::from("true"), String::from("0"), String::from("0"), String::from("0")]),
    test_unsigned_array_sum => (r##"
        fun main() {
            val u = uintArrayOf(1u, 2u, 3u)
            val b = ubyteArrayOf(10u, 20u)
            var uSum = 0u
            var bSum = 0u
            for (x in u) { uSum += x }
            for (x in b) { bSum += x.toUInt() }
            println(uSum.toString())
            println(bSum.toString())
        }
    "##, vec![String::from("6"), String::from("30")]),
    test_unsigned_array_all => (r##"
        fun main() {
            val u = uintArrayOf(1u, 4u, 9u)
            val b = ubyteArrayOf(1u, 2u)
            var allPositive = true
            for (x in u) { if (x == 0u) allPositive = false }
            var hasLow = false
            for (x in b) { if (x < 2u) hasLow = true }
            println(allPositive.toString())
            println(hasLow.toString())
        }
    "##, vec![String::from("true"), String::from("true")]),
    test_unsigned_array_contains => (r##"
        fun main() {
            val u = uintArrayOf(7u, 8u, 9u)
            val b = ubyteArrayOf(1u, 2u, 3u)
            var found = false
            var missing = false
            for (x in u) { if (x == 8u) found = true }
            for (x in b) { if (x == 9u) missing = true }
            println(found.toString())
            println(missing.toString())
        }
    "##, vec![String::from("true"), String::from("false")]),
    test_unsigned_array_casts => (r##"
        fun main() {
            val u = 255u
            val b = 255u.toUByte()
            val s = 65535u.toUShort()
            val l = 1024uL
            println(u.toByte().toString())
            println(b.toInt())
            println(s.toInt())
            println(l.toInt())
        }
    "##, vec![String::from("-1"), String::from("255"), String::from("65535"), String::from("1024")]),
    test_unsigned_array_join => (r##"
        fun main() {
            val u = uintArrayOf(1u, 2u, 3u)
            val b = ubyteArrayOf(4u, 5u, 6u)
            println(u.joinToString("|"))
            println(b.joinToString("|"))
        }
    "##, vec![String::from("1|2|3"), String::from("4|5|6")]),
    test_unsigned_array_reference_behavior => (r##"
        fun main() {
            val original = uintArrayOf(1u, 2u)
            val alias = original
            alias[1] = 9u
            println(original[1].toString())
        }
    "##, vec![String::from("9")]),
    test_unsigned_array_iteration_index => (r##"
        fun main() {
            val s = ushortArrayOf(3u, 4u, 5u)
            var out = ""
            for (i in s.indices) {
                out = out + s[i].toString()
                if (i + 1 < s.size) { out = out + "," }
            }
            println(out)
        }
    "##, vec![String::from("3,4,5")]),
    test_unsigned_array_conversion_roundtrip => (r##"
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
    "##, vec![String::from("3"), String::from("3"), String::from("2,2,2")]),
    test_unsigned_array_not_equal => (r##"
        fun main() {
            val a = ulongArrayOf(1uL, 2uL)
            val b = ulongArrayOf(1uL, 3uL)
            println((a == b).toString())
        }
    "##, vec![String::from("false")]),
    test_unsigned_array_minmax => (r##"
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
    "##, vec![String::from("1"), String::from("9")]),
}

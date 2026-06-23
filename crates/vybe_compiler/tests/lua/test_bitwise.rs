//! Bitwise operators — Lua 5.3+ manual §3.4.2.

lua_print! {
    bitwise_and_masks_bits => { "print(0xF & 0x3)\n", "3" },
    bitwise_or_combines_bits => { "print(8 | 1)\n", "9" },
    bitwise_xor_toggles_bits => { "print(5 ~ 3)\n", "6" },
    bitwise_not_inverts => { "print(~0)\n", "-1" },
    left_shift_multiplies_by_power_of_two => { "print(1 << 4)\n", "16" },
    right_shift_divides_by_power_of_two => { "print(16 >> 2)\n", "4" },
    bitwise_and_precedence_over_or => { "print(1 | 2 & 4)\n", "1" },
    bitwise_xor_is_associative => { "print(1 ~ 2 ~ 3)\n", "0" },
    floor_division_is_not_bitwise_shift => { "print(8 // 2)\n", "4" },
    bitwise_not_on_negative_one => { "print(~-1)\n", "0" },
    bitwise_shift_zero_bits_is_identity => { "print(9 << 0)\n", "9" },
    bitwise_and_with_zero_clears_bits => { "print(0xFF & 0)\n", "0" },
    bitwise_or_preserves_high_bits => { "print(0xF0 | 0x0F)\n", "255" },
    bitwise_xor_self_is_zero => { "print(7 ~ 7)\n", "0" },
    left_shift_large_moves_one_bit => { "print(1 << 8)\n", "256" },
    right_shift_on_odd_truncates => { "print(7 >> 1)\n", "3" },
    bitwise_and_with_mask_extracts_low_byte => { "print(0x12FF & 0xFF)\n", "255" },
    bitwise_not_flips_all_bits_of_zero => { "print(~0 == -1)\n", "true" },
    combined_shift_and_mask_for_nibble => { "print((0xAB >> 4) & 0xF)\n", "10" },
}

//! `fixed` pinning via `&arr[0]`, `unsafe` blocks, and pointer dereference prints.
//! GAP: fixed-buffer pointer coverage is thin beyond whole-array pinning.

csharp_cases! {
    fixed_byte_ptr_from_first_element_reads_value => {
        r#"byte[] arr={65,66,67}; unsafe{fixed(byte* ptr=&arr[0]){Console.WriteLine(*ptr);}}"#,
        ["65"]
    };

    fixed_int_ptr_from_first_element_reads_value => {
        r#"int[] arr={10,20,30}; unsafe{fixed(int* ptr=&arr[0]){Console.WriteLine(*ptr);}}"#,
        ["10"]
    };

    fixed_byte_ptr_index_one_reads_second_slot => {
        r#"byte[] arr={1,2,3}; unsafe{fixed(byte* ptr=&arr[0]){Console.WriteLine(ptr[1]);}}"#,
        ["2"]
    };

    fixed_int_ptr_arithmetic_reads_third_element => {
        r#"int[] arr={5,6,7}; unsafe{fixed(int* ptr=&arr[0]){Console.WriteLine(*(ptr+2));}}"#,
        ["7"]
    };

    fixed_byte_ptr_write_updates_array_element => {
        r#"byte[] arr={1,2,3}; unsafe{fixed(byte* ptr=&arr[0]){ptr[1]=99;}} Console.WriteLine(arr[1]);"#,
        ["99"]
    };

    fixed_int_ptr_dereference_write_mutates_source => {
        r#"int[] arr={4,5,6}; unsafe{fixed(int* ptr=&arr[0]){*(ptr+1)=88;}} Console.WriteLine(arr[1]);"#,
        ["88"]
    };

    fixed_char_ptr_reads_first_character_code => {
        r#"char[] arr={'A','B'}; unsafe{fixed(char* ptr=&arr[0]){Console.WriteLine(*ptr);}}"#,
        ["65"]
    };

    fixed_short_ptr_reads_negative_value => {
        r#"short[] arr={-3,4}; unsafe{fixed(short* ptr=&arr[0]){Console.WriteLine(*ptr);}}"#,
        ["-3"]
    };

    fixed_long_ptr_reads_large_literal => {
        r#"long[] arr={10000000000L,2L}; unsafe{fixed(long* ptr=&arr[0]){Console.WriteLine(*ptr>0);}}"#,
        ["True"]
    };

    fixed_float_ptr_reads_fractional_value => {
        r#"float[] arr={1.5f,2.5f}; unsafe{fixed(float* ptr=&arr[0]){Console.WriteLine(*ptr==1.5f);}}"#,
        ["True"]
    };

    fixed_double_ptr_reads_second_element => {
        r#"double[] arr={1.1,2.2}; unsafe{fixed(double* ptr=&arr[0]){Console.WriteLine(*(ptr+1));}}"#,
        ["2.2"]
    };

    fixed_bool_ptr_reads_true_literal => {
        r#"bool[] arr={true,false}; unsafe{fixed(bool* ptr=&arr[0]){Console.WriteLine(*ptr);}}"#,
        ["True"]
    };

    fixed_byte_ptr_from_middle_index_reads_offset => {
        r#"byte[] arr={10,20,30,40}; unsafe{fixed(byte* ptr=&arr[2]){Console.WriteLine(*ptr);}}"#,
        ["30"]
    };

    fixed_int_ptr_subtract_returns_to_first_element => {
        r#"int[] arr={11,22,33}; unsafe{fixed(int* ptr=&arr[1]){Console.WriteLine(*(ptr-1));}}"#,
        ["11"]
    };

    fixed_byte_ptr_increment_moves_to_next_byte => {
        r#"byte[] arr={7,8,9}; unsafe{fixed(byte* ptr=&arr[0]){ptr++; Console.WriteLine(*ptr);}}"#,
        ["8"]
    };

    fixed_int_ptr_decrement_reads_previous_slot => {
        r#"int[] arr={3,4,5}; unsafe{fixed(int* ptr=&arr[2]){ptr--; Console.WriteLine(*ptr);}}"#,
        ["4"]
    };

    fixed_byte_ptr_null_check_is_false_for_valid_pin => {
        r#"byte[] arr={1}; unsafe{fixed(byte* ptr=&arr[0]){Console.WriteLine(ptr==null);}}"#,
        ["False"]
    };

    fixed_int_ptr_distance_between_elements_is_one => {
        r#"int[] arr={1,2,3}; unsafe{fixed(int* a=&arr[0]){fixed(int* b=&arr[1]){Console.WriteLine(b-a);}}}"#,
        ["1"]
    };

    fixed_byte_ptr_zero_length_array_still_pins_first => {
        r#"byte[] arr={42}; unsafe{fixed(byte* ptr=&arr[0]){Console.WriteLine(*ptr);}}"#,
        ["42"]
    };

    fixed_int_ptr_from_last_index_reads_tail => {
        r#"int[] arr={100,200,300}; unsafe{fixed(int* ptr=&arr[2]){Console.WriteLine(*ptr);}}"#,
        ["300"]
    };

    fixed_byte_ptr_write_first_element_changes_array => {
        r#"byte[] arr={1,2}; unsafe{fixed(byte* ptr=&arr[0]){*ptr=77;}} Console.WriteLine(arr[0]);"#,
        ["77"]
    };

    fixed_int_ptr_sum_via_loop_accumulates => {
        r#"int[] arr={1,2,3,4}; int sum=0; unsafe{fixed(int* ptr=&arr[0]){for(int i=0;i<4;i++){sum+=ptr[i];}}} Console.WriteLine(sum);"#,
        ["10"]
    };

    fixed_byte_ptr_copy_between_offsets => {
        r#"byte[] arr={1,2,3}; unsafe{fixed(byte* ptr=&arr[0]){ptr[2]=ptr[0];}} Console.WriteLine(arr[2]);"#,
        ["1"]
    };

    fixed_int_ptr_nested_unsafe_reads_value => {
        r#"int[] arr={9}; unsafe{fixed(int* outer=&arr[0]){unsafe{Console.WriteLine(*outer);}}}"#,
        ["9"]
    };

    fixed_byte_ptr_from_local_array_reads_backing_store => {
        r#"unsafe{byte[] arr={5,6}; fixed(byte* ptr=&arr[0]){Console.WriteLine(ptr[1]);}}"#,
        ["6"]
    };

    fixed_int_ptr_post_increment_reads_original_then_advances => {
        r#"int[] arr={2,3}; unsafe{fixed(int* ptr=&arr[0]){Console.WriteLine(*ptr++); Console.WriteLine(*ptr);}}"#,
        ["2", "3"]
    };

    fixed_byte_ptr_pre_increment_advances_before_read => {
        r#"byte[] arr={1,2}; unsafe{fixed(byte* ptr=&arr[0]){Console.WriteLine(*++ptr);}}"#,
        ["2"]
    };

    fixed_int_ptr_address_of_element_equals_base_plus_offset => {
        r#"int[] arr={10,20,30}; unsafe{fixed(int* basePtr=&arr[0]){fixed(int* off=&arr[2]){Console.WriteLine(off-basePtr);}}}"#,
        ["2"]
    };

    fixed_sbyte_ptr_reads_signed_byte => {
        r#"sbyte[] arr={-1,1}; unsafe{fixed(sbyte* ptr=&arr[0]){Console.WriteLine(*ptr);}}"#,
        ["-1"]
    };

    fixed_ushort_ptr_reads_unsigned_short => {
        r#"ushort[] arr={65000,1}; unsafe{fixed(ushort* ptr=&arr[0]){Console.WriteLine(*ptr);}}"#,
        ["65000"]
    };

    fixed_uint_ptr_reads_unsigned_int => {
        r#"uint[] arr={3000000000u,1u}; unsafe{fixed(uint* ptr=&arr[0]){Console.WriteLine(*ptr>0);}}"#,
        ["True"]
    };

    fixed_ulong_ptr_reads_unsigned_long => {
        r#"ulong[] arr={18446744073709551615UL,0UL}; unsafe{fixed(ulong* ptr=&arr[1]){Console.WriteLine(*ptr);}}"#,
        ["0"]
    };

    fixed_byte_ptr_from_field_backed_array => {
        r#"class Holder{public byte[] Data={9,8};} var h=new Holder(); unsafe{fixed(byte* ptr=&h.Data[0]){Console.WriteLine(*ptr);}}"#,
        ["9"]
    };

    fixed_int_ptr_method_local_array => {
        r#"int Read(){int[] arr={42}; unsafe{fixed(int* ptr=&arr[0]){return *ptr;}} return 0;} Console.WriteLine(Read());"#,
        ["42"]
    };

    fixed_byte_ptr_two_fixed_blocks_same_array => {
        r#"byte[] arr={1,2,3}; unsafe{fixed(byte* a=&arr[0]){fixed(byte* b=&arr[1]){Console.WriteLine(*a+b[0]);}}}"#,
        ["3"]
    };

    fixed_int_ptr_compare_elements_via_pointers => {
        r#"int[] arr={5,5,9}; unsafe{fixed(int* ptr=&arr[0]){Console.WriteLine(ptr[0]==ptr[1]); Console.WriteLine(ptr[0]==ptr[2]);}}"#,
        ["True", "False"]
    };

    fixed_byte_ptr_xor_toggle_bit_in_place => {
        r#"byte[] arr={0b1010}; unsafe{fixed(byte* ptr=&arr[0]){*ptr=(byte)(*ptr^0b1111);}} Console.WriteLine(arr[0]);"#,
        ["5"]
    };

    fixed_int_ptr_readonly_span_style_index_from_end => {
        r#"int[] arr={1,2,3,4}; unsafe{fixed(int* ptr=&arr[0]){Console.WriteLine(ptr[3]);}}"#,
        ["4"]
    };

    fixed_byte_ptr_assign_from_stack_value => {
        r#"byte[] arr={0}; byte temp=33; unsafe{fixed(byte* ptr=&arr[0]){*ptr=temp;}} Console.WriteLine(arr[0]);"#,
        ["33"]
    };

    fixed_int_ptr_swap_two_elements => {
        r#"int[] arr={1,9}; unsafe{fixed(int* ptr=&arr[0]){int t=ptr[0]; ptr[0]=ptr[1]; ptr[1]=t;}} Console.WriteLine(arr[0]); Console.WriteLine(arr[1]);"#,
        ["9", "1"]
    };

    fixed_byte_ptr_read_after_array_reassign_same_buffer => {
        r#"byte[] arr={7,8}; unsafe{fixed(byte* ptr=&arr[0]){Console.WriteLine(ptr[1]); arr[1]=55; Console.WriteLine(ptr[1]);}}"#,
        ["8", "55"]
    };

    fixed_int_ptr_nullable_struct_false_reads_zero => {
        r#"int[] arr={0,1}; unsafe{fixed(int* ptr=&arr[0]){Console.WriteLine(*ptr==0);}}"#,
        ["True"]
    };

    fixed_char_ptr_write_updates_string_builder_backing => {
        r#"char[] arr={'x','y'}; unsafe{fixed(char* ptr=&arr[0]){ptr[1]='z';}} Console.WriteLine(arr[1]);"#,
        ["122"]
    };

    fixed_int_ptr_pointer_equality_same_element => {
        r#"int[] arr={1,2}; unsafe{fixed(int* a=&arr[0]){fixed(int* b=&arr[0]){Console.WriteLine(a==b);}}}"#,
        ["True"]
    };

    fixed_byte_ptr_pointer_inequality_different_indices => {
        r#"byte[] arr={1,2}; unsafe{fixed(byte* a=&arr[0]){fixed(byte* b=&arr[1]){Console.WriteLine(a==b);}}}"#,
        ["False"]
    };

    fixed_int_ptr_scale_offset_by_element_size => {
        r#"int[] arr={10,20,30,40}; unsafe{fixed(int* ptr=&arr[0]){Console.WriteLine(*(ptr+3));}}"#,
        ["40"]
    };

    fixed_byte_ptr_read_third_from_one_based_style_offset => {
        r#"byte[] arr={2,4,6,8}; unsafe{fixed(byte* ptr=&arr[0]){Console.WriteLine(ptr[2]);}}"#,
        ["6"]
    };

    fixed_int_ptr_clear_slot_via_dereference => {
        r#"int[] arr={5,6,7}; unsafe{fixed(int* ptr=&arr[0]){*(ptr+1)=0;}} Console.WriteLine(arr[1]);"#,
        ["0"]
    };

    fixed_byte_ptr_max_byte_value_roundtrip => {
        r#"byte[] arr={255,0}; unsafe{fixed(byte* ptr=&arr[0]){Console.WriteLine(*ptr);}}"#,
        ["255"]
    };

    fixed_int_ptr_min_int_value_roundtrip => {
        r#"int[] arr={-2147483648,1}; unsafe{fixed(int* ptr=&arr[0]){Console.WriteLine(*ptr<0);}}"#,
        ["True"]
    };
}

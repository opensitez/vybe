use super::helpers::compile_ok;

macro_rules! c { ($name:ident, $src:expr) => { #[test] fn $name() { compile_ok($src); } }; }

c!(scr_print_01, "program t\nprint *, 'hello'\nend program t\n");
c!(scr_write_stdout_02, "program t\nwrite(*,*) 'hello'\nend program t\n");
c!(scr_read_stdin_03, "program t\ninteger :: x\nread(*,*) x\nend program t\n");
c!(scr_format_04, "program t\ninteger :: x=1\nwrite(*,'(I3)') x\nend program t\n");
c!(scr_advance_no_05, "program t\nwrite(*,'(A)',advance='no') 'x'\nend program t\n");
c!(scr_tab_06, "program t\nwrite(*,'(T5,A)') 'x'\nend program t\n");
c!(scr_slash_07, "program t\nwrite(*,'(/,A)') 'x'\nend program t\n");
c!(scr_repeat_08, "program t\ninteger :: x=1\nwrite(*,'(3I2)') x,x,x\nend program t\n");
c!(scr_sign_09, "program t\ninteger :: x=1\nwrite(*,'(SP,I3)') x\nend program t\n");
c!(scr_blank_10, "program t\ninteger :: x=1\nwrite(*,'(BN,I3)') x\nend program t\n");
c!(scr_round_11, "program t\nreal :: x=1.2\nwrite(*,'(F5.1,ROUND=\"UP\")') x\nend program t\n");
c!(scr_decimal_12, "program t\nreal :: x=1.2\nwrite(*,'(F5.1,DECIMAL=\"POINT\")') x\nend program t\n");
c!(scr_char_13, "program t\ncharacter(len=5) :: s='abc'\nprint *, trim(s)\nend program t\n");
c!(scr_integer_14, "program t\ninteger :: i=10\nprint *, i\nend program t\n");
c!(scr_real_15, "program t\nreal :: r=1.5\nprint *, r\nend program t\n");
c!(scr_complex_16, "program t\ncomplex :: z=(1.0,2.0)\nprint *, z\nend program t\n");
c!(scr_logical_17, "program t\nlogical :: l=.true.\nprint *, l\nend program t\n");
c!(scr_array_18, "program t\ninteger :: a(3)=[1,2,3]\nprint *, a\nend program t\n");
c!(scr_namelist_19, "program t\ninteger :: x=1\nnamelist /g/ x\nwrite(*,nml=g)\nend program t\n");
c!(scr_internal_write_20, "program t\ncharacter(len=20) :: buf\nwrite(buf,*) 42\nprint *, trim(buf)\nend program t\n");
c!(scr_internal_read_21, "program t\ncharacter(len=20) :: buf='42'\ninteger :: x\nread(buf,*) x\nprint *, x\nend program t\n");
c!(scr_prompt_22, "program t\nwrite(*,'(A)',advance='no') 'Enter:'\nend program t\n");
c!(scr_multi_line_23, "program t\nprint *, 'a'\nprint *, 'b'\nend program t\n");
c!(scr_format_label_24, "program t\ninteger :: x=3\nwrite(*,100) x\n100 format(I3)\nend program t\n");
c!(scr_carriage_25, "program t\nwrite(*,'(A,/,A)') 'a','b'\nend program t\n");
c!(scr_list_directed_26, "program t\ninteger :: a=1,b=2\nwrite(*,*) a,b\nend program t\n");
c!(scr_scale_27, "program t\nreal :: x=1.23\nwrite(*,'(1P,E10.2)') x\nend program t\n");
c!(scr_colon_28, "program t\ninteger :: x=1\nwrite(*,'(I2,:,I2)') x\nend program t\n");
c!(scr_position_29, "program t\nwrite(*,'(X,A)') 'x'\nend program t\n");
c!(scr_string_concat_30, "program t\nprint *, 'a'//'b'\nend program t\n");
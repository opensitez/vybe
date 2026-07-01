use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(fmt_i_01,"program p
integer :: x=1
write(*,'(I3)') x
end program p
");
c!(fmt_f_02,"program p
real :: x=1.5
write(*,'(F5.1)') x
end program p
");
c!(fmt_e_03,"program p
real :: x=1.5
write(*,'(E10.2)') x
end program p
");
c!(fmt_g_04,"program p
real :: x=1.5
write(*,'(G10.2)') x
end program p
");
c!(fmt_a_05,"program p
character(len=3)::s='abc'
write(*,'(A)') s
end program p
");
c!(fmt_l_06,"program p
logical::x=.true.
write(*,'(L3)') x
end program p
");
c!(fmt_repeat_07,"program p
integer::x=1
write(*,'(3I2)') x,x,x
end program p
");
c!(fmt_pos_08,"program p
write(*,'(T5,A)') 'x'
end program p
");
c!(fmt_colon_09,"program p
integer::x=1
write(*,'(I2,:,I2)') x
end program p
");
c!(fmt_slash_10,"program p
write(*,'(/,A)') 'x'
end program p
");
c!(fmt_scale_11,"program p
real::x=1.23
write(*,'(1P,E10.2)') x
end program p
");
c!(fmt_sign_12,"program p
integer::x=1
write(*,'(SP,I3)') x
end program p
");
c!(fmt_blank_13,"program p
integer::x=1
write(*,'(BN,I3)') x
end program p
");
c!(fmt_round_14,"program p
real::x=1.2
write(*,'(F5.1,ROUND=\"UP\")') x
end program p
");
c!(fmt_decimal_15,"program p
real::x=1.2
write(*,'(F5.1,DECIMAL=\"POINT\")') x
end program p
");
c!(fmt_internal_16,"program p
character(len=20)::buf
write(buf,'(I3)') 42
print *, trim(buf)
end program p
");
c!(fmt_read_17,"program p
character(len=20)::buf='42'
integer::x
read(buf,'(I2)') x
print *,x
end program p
");
c!(fmt_list_18,"program p
integer::a=1,b=2
write(*,*) a,b
end program p
");
c!(fmt_adv_19,"program p
write(*,'(A)',advance='no') 'x'
end program p
");
c!(fmt_label_20,"program p
integer::x=1
write(*,100) x
100 format(I3)
end program p
");
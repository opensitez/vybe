use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(ctrl_associate_01,"program p
integer::x=1
associate(y=>x)
 print *, y
end associate
end program p
");
c!(ctrl_block_02,"program p
block
 integer::x
 x=1
 print *,x
end block
end program p
");
c!(ctrl_change_team_03,"program p
change team (team=1)
end team
end program p
");
c!(ctrl_critical_04,"program p
critical
 print *,1
end critical
end program p
");
c!(ctrl_error_stop_05,"program p
error stop
end program p
");
c!(ctrl_stop_code_06,"program p
stop 1
end program p
");
c!(ctrl_cycle_07,"program p
integer::i
do i=1,3
 cycle
end do
end program p
");
c!(ctrl_exit_08,"program p
integer::i
do i=1,3
 exit
end do
end program p
");
c!(ctrl_return_09,"subroutine s()
return
end subroutine s
");
c!(ctrl_goto_computed_10,"program p
integer::k=1
go to (10,20), k
10 continue
20 continue
end program p
");
c!(ctrl_goto_assigned_11,"program p
integer :: n
assign 10 to n
go to n
10 continue
end program p
");
c!(ctrl_arith_if_12,"program p
integer::x=1
if (x) 10,20,30
10 continue
20 continue
30 continue
end program p
");
c!(ctrl_block_if_13,"program p
integer::x=1
if (x==1) then
 print *,1
else
 print *,2
end if
end program p
");
c!(ctrl_select_case_14,"program p
integer::x=1
select case(x)
 case(1)
  print *,1
end select
end program p
");
c!(ctrl_do_15,"program p
integer::i
do i=1,3
 print *,i
end do
end program p
");
c!(ctrl_do_while_16,"program p
integer::i=0
do while(i<3)
 i=i+1
end do
end program p
");
c!(ctrl_do_concurrent_17,"program p
integer::i
do concurrent (i=1:3)
 print *,i
end do
end program p
");
c!(ctrl_where_18,"program p
integer::a(3)=[1,2,3]
where(a>1) a=a+1
print *,a
end program p
");
c!(ctrl_forall_19,"program p
integer::a(3)
forall(i=1:3) a(i)=i
print *,a
end program p
");
c!(ctrl_select_type_20,"program p
class(*),allocatable::x
allocate(integer::x)
select type(x)
 type is(integer)
  print *,x
 class default
end select
end program p
");
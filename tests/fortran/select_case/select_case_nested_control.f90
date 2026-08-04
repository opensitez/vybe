! vybe-test: fortran/select_case/select_case_nested_control
! origin: languages/fortran/tests/fortran/test_select_case.rs
program t
integer :: x
integer :: y
x = 3
y = 0
select case (x)
case (1:2)
 y = 1
case (3)
 select case (x)
 case (3)
  y = 3
 case default
  y = 9
 end select
case default
 y = 5
end select
if ((y) /= 3) then
    print *, "FAIL: want [3] got [", y, "]"
    stop 1
end if
end program t

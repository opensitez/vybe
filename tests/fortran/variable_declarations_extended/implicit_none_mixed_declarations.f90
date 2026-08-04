! vybe-test: fortran/variable_declarations_extended/implicit_none_mixed_declarations
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer :: a = 1
real :: b = 2.0
logical :: c = .true.
character(len=1) :: d = "x"
if ((a) /= 1) then
    print *, "FAIL: want [1] got [", a, "]"
    stop 1
end if
if ((b) /= 2) then
    print *, "FAIL: want [2] got [", b, "]"
    stop 1
end if
if ((c) .neqv. .true.) then
    print *, "FAIL: want [true] got [", c, "]"
    stop 1
end if
if (trim(d) /= "x") then
    print *, "FAIL: want [x] got [", d, "]"
    stop 1
end if
end program t

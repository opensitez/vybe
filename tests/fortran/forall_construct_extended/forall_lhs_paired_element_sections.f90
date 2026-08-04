! vybe-test: fortran/forall_construct_extended/forall_lhs_paired_element_sections
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: a(8)
a = 0
forall (i = 1:4)
a(2 * i - 1:2 * i) = i
end forall
if ((a(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1), "]"
    stop 1
end if
if ((a(2)) /= 1) then
    print *, "FAIL: want [1] got [", a(2), "]"
    stop 1
end if
if ((a(7)) /= 4) then
    print *, "FAIL: want [4] got [", a(7), "]"
    stop 1
end if
if ((a(8)) /= 4) then
    print *, "FAIL: want [4] got [", a(8), "]"
    stop 1
end if
end program t

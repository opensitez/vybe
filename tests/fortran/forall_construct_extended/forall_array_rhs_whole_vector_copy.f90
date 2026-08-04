! vybe-test: fortran/forall_construct_extended/forall_array_rhs_whole_vector_copy
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: src(4) = [10, 20, 30, 40]
integer :: dst(4) = 0
forall (i = 1:1)
dst(1:4) = src(1:4)
end forall
if ((dst(1)) /= 10) then
    print *, "FAIL: want [10] got [", dst(1), "]"
    stop 1
end if
if ((dst(3)) /= 30) then
    print *, "FAIL: want [30] got [", dst(3), "]"
    stop 1
end if
if ((sum(dst)) /= 100) then
    print *, "FAIL: want [100] got [", sum(dst), "]"
    stop 1
end if
end program t

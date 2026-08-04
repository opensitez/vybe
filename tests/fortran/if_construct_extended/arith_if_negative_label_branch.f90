! vybe-test: fortran/if_construct_extended/arith_if_negative_label_branch
! origin: languages/fortran/tests/fortran/test_if_construct_extended.rs
program t
real :: x = -2.5
if (x) 10, 20, 30
10 print *, "negative"; goto 99
20 print *, "zero"; goto 99
30 print *, "positive"
99 continue
end program t

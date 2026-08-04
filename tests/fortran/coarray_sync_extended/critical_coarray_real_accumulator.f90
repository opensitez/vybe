! vybe-test: fortran/coarray_sync_extended/critical_coarray_real_accumulator
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
real :: total[*]
total = 0.0
sync all
critical
total[1] = total[1] + real(this_image())
end critical
sync all
if (this_image() == 1) print *, total
end program t

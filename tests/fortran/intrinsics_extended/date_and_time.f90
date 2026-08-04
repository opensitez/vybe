! vybe-test: fortran/intrinsics_extended/date_and_time
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs

program test
    character(len=8) :: d
    character(len=10) :: t
    call date_and_time(date=d, time=t)
    print *, "got date"
end program test

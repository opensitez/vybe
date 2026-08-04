! vybe-test: fortran/fortran2018_extended/reduce_logical_or_chain
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
logical :: flags(3) = [.false., .true., .false.]
if ((reduce(flags, operator(.or.))) .neqv. .true.) then
    print *, "FAIL: want [true] got [", reduce(flags, operator(.or.)), "]"
    stop 1
end if
end program t

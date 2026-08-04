! vybe-test: fortran/legacy_data_extended/common_logical_pair_and
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
logical :: f1, f2
common /flags/ f1, f2
f1 = .true.
f2 = .false.
if ((f1 .and. f2) .neqv. .false.) then
    print *, "FAIL: want [false] got [", f1 .and. f2, "]"
    stop 1
end if
if ((f1 .or. f2) .neqv. .true.) then
    print *, "FAIL: want [true] got [", f1 .or. f2, "]"
    stop 1
end if
end program t

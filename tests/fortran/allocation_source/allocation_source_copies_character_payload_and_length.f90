! vybe-test: fortran/allocation_source/allocation_source_copies_character_payload_and_length
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program t
character(len=:), allocatable :: s
allocate(character(len=3) :: s, source='abc')
if ((len(s)) /= 3) then
    print *, "FAIL: want [3] got [", len(s), "]"
    stop 1
end if
if (trim(s) /= "abc") then
    print *, "FAIL: want [abc] got [", s, "]"
    stop 1
end if
end program t

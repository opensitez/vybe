! vybe-test: fortran/pointer_alloc_extended/alloc_char_two_element_trim
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
character(len=4), allocatable :: words(:)
allocate(words(2))
words(1) = 'ab'
words(2) = 'cd'
if (trim(trim(words(1))) /= "ab") then
    print *, "FAIL: want [ab] got [", trim(words(1)), "]"
    stop 1
end if
if ((len_trim(words(2))) /= 2) then
    print *, "FAIL: want [2] got [", len_trim(words(2)), "]"
    stop 1
end if
deallocate(words)
end program t

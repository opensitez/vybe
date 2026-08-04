! vybe-test: fortran/character_arrays/char_array_function_result
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=5) :: words(3)
    integer :: longest
    words = ['hi   ', 'hello', 'hey  ']
    longest = find_longest(words)
    print *, longest
contains
    function find_longest(arr) result(maxlen)
        character(len=*), intent(in) :: arr(:)
        integer :: maxlen, i
        maxlen = 0
        do i = 1, size(arr)
            maxlen = max(maxlen, len_trim(arr(i)))
        end do
    end function find_longest
end program test

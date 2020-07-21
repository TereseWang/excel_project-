#demonstrate how to do list comprehension w conditional

print([i for i in range (10) if i % 2 == 0])

#will not get an error

ls = ['0', '1', '2']

upd_ls = [i if i != '0' else 'none' for i in ls]

print(upd_ls)

#if there is a letter in the ls there will be a valueerror

ls1 = ['0', 'k', '1']

# try:
# 	upd_ls1 = [i if i != '0' else 'none' for i in ls]
# except ValueError:
# 	print('there is a letter')

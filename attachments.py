# -*- coding: utf-8 -*-#

# -----------------------------------------------------------
# Name:         attachments
# Author:       Jerem
# Date:         2022/8/21
# Description:  
# -----------------------------------------------------------

import sys
import os
from tqdm import tqdm
import shutil

wiz_dir = ''
while not os.path.exists(wiz_dir):
    wiz_dir = input('please input wiz_path like: ../{account}/:')

store_path = input('please input target_path, end with "/":')
if store_path[-1] == '/':
    store_path = store_path[:-1]
if not os.path.exists(store_path):
    os.makedirs(store_path)

level = int(input('please input the restored architecture, level0-"/files", level1="/arch1/files", level2="/arch1/arch2/files":'))

def count(source_dir):
    n_folder, n_ziw, n_other = 1, 0, 0
    for f in os.listdir(wiz_dir + source_dir):
        if os.path.isdir(wiz_dir + source_dir + '/' + f):
            nums = count(source_dir + '/' + f)
            n_folder += nums[0]
            n_ziw += nums[1]
            n_other += nums[2]
        else:
            if f[-4:] == '.ziw':
                n_ziw += 1
            else:
                n_other += 1
    return n_folder, n_ziw, n_other

nums = count('')
print('n_folder, n_ziw, n_file_to_copy = ', nums)

fail_list = []
with tqdm(total=nums[2]) as pbar:
    def dfs(source_dir):
        for f in os.listdir(wiz_dir + source_dir):
            if os.path.isdir(wiz_dir + source_dir + '/' + f):
                dfs(source_dir + '/' + f)
            elif f[-4:] != '.ziw':
                try:
                    if level == 0:
                        target_file = store_path + '/' + f
                    else:
                        target_file = store_path + '/' + '/'.join(source_dir.split('/')[:level])
                        if not os.path.exists(target_file):
                            os.makedirs(target_file)
                        target_file += '/' + f
                    shutil.copy(wiz_dir + source_dir + '/' + f, target_file)
                    pbar.update(1)
                except:
                    fail_list.append(wiz_dir + source_dir + '/' + f)
    dfs('')

print('Finish! num of fail: ', len(fail_list))
print('print the list of fail?y/n')
if input('y/n?').lower() == 'y':
    for f in fail_list:
        print(f)






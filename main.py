import os
import sys
import subprocess
import tempfile
from pathlib import Path
from tkinter import *
from tkinter import ttk, messagebox, filedialog
import shutil
import re

class EnvEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("环境变量编辑器")
        self.root.geometry("900x650")
        
        self.user_env_vars = {}
        self.system_env_vars = {}
        
        self.create_widgets()
        self.load_user_env_vars()
        self.load_system_env_vars()
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(W, E, N, S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        title_label = ttk.Label(main_frame, text="环境变量编辑器", font=("Microsoft YaHei", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        user_frame = ttk.LabelFrame(main_frame, text="用户环境变量", padding="10")
        user_frame.grid(row=1, column=0, columnspan=2, sticky=(W, E, N, S), pady=5)
        
        self.user_tree = ttk.Treeview(user_frame, columns=("key", "value"), show="headings", height=8)
        self.user_tree.heading("key", text="变量名")
        self.user_tree.heading("value", text="值")
        self.user_tree.column("key", width=200)
        self.user_tree.column("value", width=400)
        self.user_tree.grid(row=0, column=0, columnspan=4, sticky=(W, E, N, S), pady=5)
        
        user_btn_frame = ttk.Frame(user_frame)
        user_btn_frame.grid(row=1, column=0, columnspan=4, pady=5)
        ttk.Button(user_btn_frame, text="添加", command=self.add_user_var).pack(side=LEFT, padx=5)
        ttk.Button(user_btn_frame, text="修改", command=self.edit_user_var).pack(side=LEFT, padx=5)
        ttk.Button(user_btn_frame, text="删除", command=self.delete_user_var).pack(side=LEFT, padx=5)
        ttk.Button(user_btn_frame, text="刷新", command=self.load_user_env_vars).pack(side=LEFT, padx=5)
        
        system_frame = ttk.LabelFrame(main_frame, text="系统环境变量", padding="10")
        system_frame.grid(row=2, column=0, columnspan=2, sticky=(W, E, N, S), pady=5)
        
        self.system_tree = ttk.Treeview(system_frame, columns=("key", "value"), show="headings", height=8)
        self.system_tree.heading("key", text="变量名")
        self.system_tree.heading("value", text="值")
        self.system_tree.column("key", width=200)
        self.system_tree.column("value", width=400)
        self.system_tree.grid(row=0, column=0, columnspan=4, sticky=(W, E, N, S), pady=5)
        
        system_btn_frame = ttk.Frame(system_frame)
        system_btn_frame.grid(row=1, column=0, columnspan=4, pady=5)
        ttk.Button(system_btn_frame, text="添加", command=self.add_system_var).pack(side=LEFT, padx=5)
        ttk.Button(system_btn_frame, text="修改", command=self.edit_system_var).pack(side=LEFT, padx=5)
        ttk.Button(system_btn_frame, text="删除", command=self.delete_system_var).pack(side=LEFT, padx=5)
        ttk.Button(system_btn_frame, text="刷新", command=self.load_system_env_vars).pack(side=LEFT, padx=5)
        
        ttk.Label(main_frame, text="注: 修改系统环境变量需要管理员权限", foreground="red").grid(row=3, column=0, columnspan=2, pady=5)
        
        main_frame.columnconfigure(0, weight=1)
        user_frame.columnconfigure(0, weight=1)
        system_frame.columnconfigure(0, weight=1)
    
    def load_user_env_vars(self):
        for item in self.user_tree.get_children():
            self.user_tree.delete(item)
        
        self.user_env_vars = {}
        
        try:
            if sys.platform == 'win32':
                result = subprocess.run(
                    ['powershell', '-Command', '[Environment]::GetEnvironmentVariables("User") | ConvertTo-Json -Depth 1'],
                    capture_output=True, text=True, encoding='utf-8', errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW
                )
                if result.returncode == 0 and result.stdout.strip():
                    import json
                    env_vars = json.loads(result.stdout)
                    for key, value in env_vars.items():
                        self.user_env_vars[key] = value
                        self.user_tree.insert("", END, values=(key, value))
            else:
                for key, value in os.environ.items():
                    self.user_env_vars[key] = value
                    self.user_tree.insert("", END, values=(key, value))
        except Exception as e:
            messagebox.showerror("错误", f"加载用户环境变量失败: {str(e)}")
    
    def load_system_env_vars(self):
        for item in self.system_tree.get_children():
            self.system_tree.delete(item)
        
        self.system_env_vars = {}
        
        try:
            if sys.platform == 'win32':
                result = subprocess.run(
                    ['powershell', '-Command', '[Environment]::GetEnvironmentVariables("Machine") | ConvertTo-Json -Depth 1'],
                    capture_output=True, text=True, encoding='utf-8', errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW
                )
                if result.returncode == 0 and result.stdout.strip():
                    import json
                    env_vars = json.loads(result.stdout)
                    for key, value in env_vars.items():
                        self.system_env_vars[key] = value
                        self.system_tree.insert("", END, values=(key, value))
            else:
                for key, value in os.environ.items():
                    self.system_env_vars[key] = value
                    self.system_tree.insert("", END, values=(key, value))
        except Exception as e:
            messagebox.showerror("错误", f"加载系统环境变量失败: {str(e)}")
    
    def add_user_var(self):
        self.open_var_dialog("添加用户变量", "", "", self.user_tree, self.user_env_vars, is_user=True)
    
    def edit_user_var(self):
        selected = self.user_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要修改的变量")
            return
        item = self.user_tree.item(selected[0])
        key, value = item['values']
        self.open_var_dialog("修改用户变量", key, value, self.user_tree, self.user_env_vars, is_user=True)
    
    def delete_user_var(self):
        selected = self.user_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的变量")
            return
        item = self.user_tree.item(selected[0])
        key = item['values'][0]
        if messagebox.askyesno("确认", f"确定要删除用户变量 '{key}' 吗？"):
            try:
                if sys.platform == 'win32':
                    subprocess.run(
                        ['powershell', '-Command', f'[Environment]::SetEnvironmentVariable("{key}", $null, "User")'],
                        capture_output=True, encoding='utf-8', errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW
                    )
                del self.user_env_vars[key]
                self.user_tree.delete(selected[0])
                messagebox.showinfo("成功", "用户环境变量已删除")
            except Exception as e:
                messagebox.showerror("错误", f"删除失败: {str(e)}")
    
    def add_system_var(self):
        self.open_var_dialog("添加系统变量", "", "", self.system_tree, self.system_env_vars, is_user=False)
    
    def edit_system_var(self):
        selected = self.system_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要修改的变量")
            return
        item = self.system_tree.item(selected[0])
        key, value = item['values']
        self.open_var_dialog("修改系统变量", key, value, self.system_tree, self.system_env_vars, is_user=False)
    
    def delete_system_var(self):
        selected = self.system_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的变量")
            return
        item = self.system_tree.item(selected[0])
        key = item['values'][0]
        if messagebox.askyesno("确认", f"确定要删除系统变量 '{key}' 吗？"):
            try:
                if sys.platform == 'win32':
                    subprocess.run(
                        ['powershell', '-Command', f'[Environment]::SetEnvironmentVariable("{key}", $null, "Machine")'],
                        capture_output=True, encoding='utf-8', errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW
                    )
                del self.system_env_vars[key]
                self.system_tree.delete(selected[0])
                messagebox.showinfo("成功", "系统环境变量已删除")
            except Exception as e:
                messagebox.showerror("错误", f"删除失败: {str(e)}")
    
    def open_var_dialog(self, title, key, value, tree, env_vars, is_user):
        dialog = Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("500x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="变量名:").grid(row=0, column=0, padx=10, pady=10, sticky=W)
        key_entry = ttk.Entry(dialog, width=40)
        key_entry.grid(row=0, column=1, padx=10, pady=10)
        key_entry.insert(0, key)
        
        ttk.Label(dialog, text="变量值:").grid(row=1, column=0, padx=10, pady=10, sticky=W)
        value_text = Text(dialog, width=40, height=5)
        value_text.grid(row=1, column=1, padx=10, pady=10)
        value_text.insert("1.0", value)
        
        def save():
            new_key = key_entry.get().strip()
            new_value = value_text.get("1.0", END).strip()
            
            if not new_key:
                messagebox.showwarning("警告", "变量名不能为空")
                return
            
            try:
                if sys.platform == 'win32':
                    scope = "User" if is_user else "Machine"
                    subprocess.run(
                        ['powershell', '-Command', f'[Environment]::SetEnvironmentVariable("{new_key}", "{new_value}", "{scope}")'],
                        capture_output=True, encoding='utf-8', errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    if key and key != new_key:
                        subprocess.run(
                            ['powershell', '-Command', f'[Environment]::SetEnvironmentVariable("{key}", $null, "{scope}")'],
                            capture_output=True, encoding='utf-8', errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW
                        )
                
                if is_user:
                    self.load_user_env_vars()
                    messagebox.showinfo("成功", "用户环境变量已更新")
                else:
                    self.load_system_env_vars()
                    messagebox.showinfo("成功", "系统环境变量已更新")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {str(e)}")
                return
            
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="保存", command=save).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=10)

def main():
    root = Tk()
    app = EnvEditor(root)
    root.mainloop()

if __name__ == "__main__":
    main()

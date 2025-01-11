from design_tree_builder import build_design_tree_from_excel




def read():
    # 示例调用
    file_path = 'SJT 207.4-2018 设计文件管理制度 第4部分：设计文件的编号.xlsx'
    root_node = build_design_tree_from_excel(file_path)

    # 打印结果以检查结构
    def print_tree(node, indent=""):
        print(f"{indent}{node}")
        for sub in node.sub_levels:
            print_tree(sub, indent + "  -> ")

    print_tree(root_node)

if __name__ == "__main__":
    read()
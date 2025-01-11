import pandas as pd


def build_design_tree_from_excel(file_path):
    class DesignLevel:
        def __init__(self, level_code=None, level_name=None, parent=None, level=0):
            self.level_code = level_code  # 级别编号
            self.level_name = level_name  # 级别名字
            self.parent = parent  # 父节点
            self.level = level  # 节点等级
            self.sub_levels = []  # 子级别

        def add_sub_level(self, sub_level):
            """添加子级别"""
            if isinstance(sub_level, DesignLevel):
                self.sub_levels.append(sub_level)
            else:
                raise ValueError("sub_level must be an instance of DesignLevel")

        def __repr__(self):
            return f"DesignLevel(level={self.level}, {self.level_code}, {self.level_name})"

    def process_sub_levels(df, parent_node, start_row, end_row, level_columns=(2, 3), level=2):
        """处理给定行范围内的子节点"""
        current_level_code = None  # 当前子节点的level_code
        current_level = None  # 当前子节点对象
        next_start_row = None  # 下一级子节点的起始行号

        for i in range(start_row, end_row + 1):
            row = df.iloc[i]
            level_code_col, level_name_col = level_columns

            # 检查列索引是否在有效范围内
            if level_code_col < len(df.columns) and level_name_col < len(df.columns):
                level_code = str(row[level_code_col]).strip() if not pd.isna(row[level_code_col]) else None
                level_name = str(row[level_name_col]).strip() if not pd.isna(row[level_name_col]) else None

                # 检查是否是一个新的级别条目（非空值且与当前值不同表示新级别）
                if level_code and (level_code != current_level_code or current_level is None):
                    # 如果存在之前的 level 节点，先将其添加到父节点并处理其子节点
                    if current_level is not None:
                        parent_node.add_sub_level(current_level)
                        if next_start_row is not None:
                            next_level_columns = (level_code_col + 2, level_name_col + 2)
                            if next_level_columns[0] < len(df.columns) and next_level_columns[1] < len(df.columns):
                                process_sub_levels(df, current_level, next_start_row, i - 1, next_level_columns, level + 1)

                    # 更新当前的 level code 和起始行号
                    current_level_code = level_code
                    next_start_row = i

                    # 创建新的 level 节点
                    current_level = DesignLevel(
                        level_code=level_code,
                        level_name=level_name,
                        parent=parent_node,
                        level=level
                    )

        # 添加最后一个 level 节点，并处理其子节点
        if current_level is not None:
            parent_node.add_sub_level(current_level)
            if next_start_row is not None:
                next_level_columns = (level_code_col + 2, level_name_col + 2)
                if next_level_columns[0] < len(df.columns) and next_level_columns[1] < len(df.columns):
                    process_sub_levels(df, current_level, next_start_row, end_row, next_level_columns, level + 1)

    def read_excel_and_build_tree():
        df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)  # 假设没有标题行

        root = DesignLevel(level=0)  # 创建根节点

        current_level_1_code = None  # 当前的一级子节点的level_code
        current_level_1 = None  # 当前的一级子节点对象
        start_row = None  # 当前一级子节点的起始行号

        for index, row in df.iterrows():
            if index >= 2:  # 从第三行开始（索引为2）
                level_1_code = str(row[0]).strip() if not pd.isna(row[0]) and len(df.columns) > 0 else None
                level_1_name = str(row[1]).strip() if not pd.isna(row[1]) and len(df.columns) > 1 else None

                # 检查是否是一个新的级别1条目（非空值且与当前值不同表示新级别）
                if level_1_code and level_1_code != current_level_1_code:
                    # 如果存在之前的 level_1 节点，先将其添加到根节点，并处理其下的子节点
                    if current_level_1 is not None:
                        root.add_sub_level(current_level_1)
                        process_sub_levels(df, current_level_1, start_row, index - 1)

                    # 更新当前的 level_1 code 和起始行号
                    current_level_1_code = level_1_code
                    start_row = index

                    # 创建新的 level_1 节点
                    current_level_1 = DesignLevel(
                        level_code=level_1_code,
                        level_name=level_1_name,
                        parent=root,
                        level=1
                    )

        # 添加最后一个 level_1 节点，并处理其下的子节点
        if current_level_1 is not None:
            root.add_sub_level(current_level_1)
            process_sub_levels(df, current_level_1, start_row, len(df) - 1)

        return root

    # 构建设计文件管理树状结构
    root_node = read_excel_and_build_tree()
    return root_node

import openpyxl
import time
def get_format_dict(cell):
    #cell = ws.unmerge_cells(get_cell.coordinate)[0][0]  
    return {
        "font": cell.font,
        "alignment": cell.alignment,
        "fill": cell.fill,
        "border": cell.border,
    #"width": ws.column_dimensions[cell.column_letter].width,
       # "height": ws.row_dimensions[cell.row].height,
    }

def apply_template(target_xlsx_file, template_xlsx_file, template_sheet_name, output_name):
    template_wb = openpyxl.load_workbook(template_xlsx_file)
    target_wb = openpyxl.load_workbook(target_xlsx_file)
    
    template_ws = template_wb[template_sheet_name]
    target_ws = target_wb.active

    print(f'{template_ws.max_row=}')
    print(f'{target_ws.max_row=}')

    total_count = int( (target_ws.max_row-1) / (template_ws.max_row-1) )

    # 최초 1회만 셀 서식 저장
    format_dict = {}
    for template_row in template_ws.iter_rows(min_row=2, max_row=template_ws.max_row, max_col=template_ws.max_column):
        for template_cell in template_row:
            format_dict[template_cell.coordinate] = get_format_dict(template_cell)

    # 서식 적용
    for i in range(total_count):
        for row_target, row_template in zip(target_ws.iter_rows(min_row=i*(template_ws.max_row-1)+1, max_row=target_ws.max_row, max_col=template_ws.max_column), template_ws.iter_rows(min_row=1, max_row=template_ws.max_row, max_col=template_ws.max_column)):
            for target_cell, template_cell in zip(row_target, row_template):
                target_format = format_dict.get(template_cell.coordinate, {})
                
                if "font" in target_format:
                    target_cell.font = openpyxl.styles.Font(**target_format["font"].__dict__)

                if "alignment" in target_format:
                    target_cell.alignment = openpyxl.styles.Alignment(**target_format["alignment"].__dict__)

                if "fill" in target_format:
                    target_cell.fill = openpyxl.styles.PatternFill(**target_format["fill"].__dict__)

                if "border" in target_format:
                    target_cell.border = openpyxl.styles.Border(**target_format["border"].__dict__)

                if "width" in target_format:
                    target_cell.width = openpyxl.styles.Width(**target_format["width"].__dict__)


    # # 서식 적용
    # for target_row in target_ws.iter_rows(min_row=2, max_row=target_ws.max_row, max_col=template_ws.max_column):
    #     for target_cell, template_cell in zip(target_row, template_ws.iter_cols(min_row=1, max_row=template_ws.max_row, max_col=template_ws.max_column)):
    #         target_format = format_dict.get(template_cell.coordinate, {})
    #         target_cell.font = target_format.get("font", openpyxl.styles.Font())
    #         target_cell.alignment = target_format.get("alignment", openpyxl.styles.Alignment())
    #         target_cell.fill = target_format.get("fill", openpyxl.styles.PatternFill())
    #         target_cell.border = target_format.get("border", openpyxl.styles.Border())
    # 템플릿 문서의 열 길이와 타겟 문서의 열 길이 동일하게 조정

    # for template_col, target_col in zip(template_ws.iter_cols(), target_ws.iter_cols()):
    #     target_col[0].column_dimensions = template_col[0].column_dimensions

    #target_ws.column_dimensions['a'].width = template_ws.column_dimensions['a'].width 

    for col_letter in template_ws.iter_cols(min_row=1, max_row=1):
        col_letter = col_letter[0].column_letter
        target_ws.column_dimensions[col_letter].width = template_ws.column_dimensions[col_letter].width
    target_wb.save(output_name)

# 예제 사용법
#apply_template("1.target_xlsx_file.xlsx", "2.template_xlsx_file.xlsx")
if __name__ == "__main__" :
    target_file = fr'd:\파이썬결과물저장소\CLM\test_20240206_110231.xlsx'
    template_file = fr'd:\파이썬결과물저장소\CLM\template.xlsx'
    sheet_name = '길드도감재료'
    output_name = fr'd:\파이썬결과물저장소\CLM\test_{time.strftime("%Y%m%d_%H%M%S")}_format.xlsx'
    apply_template(target_file,template_file,sheet_name,output_name)
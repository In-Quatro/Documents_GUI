from openpyxl.styles import Border, Side, Alignment, Font


font = Font(name='Times New Roman', size=9)
font_bold = Font(name='Times New Roman', size=9, bold=True)

alignment_1 = Alignment(horizontal='center',
                      vertical='center',
                      wrap_text=True
                      )
alignment_2 = Alignment(horizontal='left',
                        vertical='center',
                        wrap_text=True
                        )
alignment_3 = Alignment(horizontal='left',
                        vertical='top',
                        wrap_text=True
                        )
border = Border(left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin'))
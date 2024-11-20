import time
from flask import Flask,Response, render_template, request, send_file, redirect, url_for, flash, abort
import pandas as pd
from io import BytesIO
import os
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Spacer
#from reportlab.pdfgen import canvas
from reportlab.platypus import Image, ListFlowable, ListItem
from reportlab.lib.units import mm
import re
from flask_wtf.csrf import CSRFProtect
from flask_wtf import FlaskForm
from werkzeug.serving import WSGIRequestHandler
from wtforms import StringField, SubmitField
from wtforms import StringField, FileField, SubmitField
from wtforms.validators import DataRequired
from werkzeug.utils import secure_filename




app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MERGED_FOLDER'] = 'merged'
#app.secret_key = 'supersecretkey'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['MERGED_FOLDER'], exist_ok=True)
app.config['SECRET_KEY'] = os.urandom(24)
#csrf = CSRFProtect(app)
from flask_talisman import Talisman

self = "'self'"
# Define a strict Content Security Policy
csp = {
    'default-src': [self],
    'script-src': [self],
    'style-src': [self],
    'img-src': [self],
    'connect-src': [self],
    'font-src': [self],
    'object-src': ["'none'"],
    'frame-ancestors': ["'none'"],
    'base-uri': [self],
    'form-action': [self],
    'upgrade-insecure-requests': []  # Forces secure requests
}

Talisman(app, content_security_policy=csp, force_https=False)


# Define a form class
class UploadForm(FlaskForm):
    name = StringField('Name', validators=[DataRequired()])
    file = FileField('File', validators=[DataRequired()])
    submit = SubmitField('Submit')

counter = 1

# 定义插入函数
def insert_column(merged_df, position):
    global counter
    new_column_name = f"id{counter}_0"
    new_column_name2 = f"column_{counter}_1"  # 第一列名称
    new_column_name3 = f"column_{counter}_2"  # 第二列名称
    # 插入新列
    merged_df.insert(position, new_column_name3, merged_df.iloc[:, 2])
    merged_df.insert(position, new_column_name2, merged_df.iloc[:, 1])
    merged_df.insert(position, new_column_name, merged_df.iloc[:, 0])

    # 计数器递增
    counter += 1



def add_page_number(canvas, doc):
    page_num = canvas.getPageNumber()
    text = f"Page {page_num}"
    canvas.setFont("Helvetica", 9)
    width, _ = letter
    canvas.drawString(650, 50, "IT Security team")
    canvas.drawCentredString(width / 2+100, 20, text)  # Position at bottom center of page


def merge_excel(users_file, details_files):
    try:
        # 读取用户文件
        users_df = pd.read_excel(users_file, dtype=str)

        # 逐步合并每一个 details 文件
        for details_file in details_files:
            details_df = pd.read_excel(details_file, dtype=str)
            users_df = pd.merge(users_df, details_df, left_on=users_df.columns[0], right_on=details_df.columns[0],
                                how='inner')

        # 对最终合并的 DataFrame 进行排序
        merged_df = users_df.sort_index()  # 排序索引
        print(merged_df)
        return merged_df
    except Exception as e:
        print(f"Error merging Excel files: {e}")
        return None


#def get_timestamped_filename(original_filename):
#    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
#    name, ext = os.path.splitext(original_filename)
#    return f"{name}_{timestamp}{ext}"

def get_timestamped_filename(filename):
    safe_filename = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', filename)  # 只允许字母、数字、下划线、破折号和点
    return f"{int(time.time())}_{safe_filename}"

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'csv', 'xlsm'}



def save_merged_file(merged_df, file_type, title, title_2, note1):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    merged_filename = f"merged_data_{timestamp}.{file_type}"
    merged_file_path = os.path.join(app.config['MERGED_FOLDER'], merged_filename)

    if file_type == 'csv':
        merged_df.to_csv(merged_file_path, index=False)
    elif file_type == 'xlsx':
        with pd.ExcelWriter(merged_file_path, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='MergedData')
    elif file_type == 'pdf':
        if merged_df.empty:
            print("The merged DataFrame is empty.")
            return None  # Early return if there's no data

        try:
            # PDF Generation
            columns = merged_df.columns.tolist()

            new_columns = []
            modified_columns = []
            for i, col in enumerate(merged_df.columns):
                if 'Unnamed' in col and i > 0:  # 确保不是第一列
                    new_columns.append(new_columns[i - 1])  # 替换为左侧相邻列名
                    modified_columns.append(i + 1)  # 记录修改列的位置（从1开始计数）
                else:
                    new_columns.append(col)

            # 设置新的列名
            merged_df.columns = new_columns

            last_in_intervals = []
            for i in range(1, len(modified_columns)):
                # 如果当前列不连续或是最后一列，记录前一列为区间末尾
                if modified_columns[i] != modified_columns[i - 1] + 1:
                    last_in_intervals.append(modified_columns[i - 1])
            # 添加最后一个区间的结尾
            last_in_intervals.append(modified_columns[-1])

            print(f"test:{last_in_intervals}")

            # 显示处理后的数据
            print("hear")
            print(merged_df)


            new_header = merged_df.columns.tolist()  # 获取原始列名
            merged_df.loc[-1] = new_header  # 在顶部插入列名作为新的一行
            merged_df.index = merged_df.index + 1  # 移动原始数据的索引

            merged_df = merged_df.sort_index()  # 排序索引

            modified_columns = []  # 用于记录被修改的列号


            doc = SimpleDocTemplate(merged_file_path, pagesize=landscape(letter),leftMargin=0*inch, rightMargin=0*inch,
            topMargin=0*inch,bottomMargin=0*inch)
            story = []
            styles = getSampleStyleSheet()


            logo_path = 'static/images/logo.jpg'  # 更新你的 logo 路径
            logo = Image(logo_path, width=2 * inch, height=1 * inch)

            # 添加右上角两行文字
            text_line1 = ""  # 更新第一行文字
            text_line2 = "Confidential"

            # 将文字分为两段
            text_paragraph1 = Paragraph(text_line1, styles['Normal'])
            text_paragraph2 = Paragraph(text_line2, styles['Normal'])

            # 使用 Spacer 调整两段文字之间的间距
            spacer = Spacer(0, 0.35 * inch)  # 调整 0.2 inch 作为间隔

            # 组合两段文字为一个元素，保持同一行
            text_block = [text_paragraph1, spacer, text_paragraph2]

            # 创建表格布局：一行两列，左边放 logo，右边放两行文字（作为一个块）
            table_data = [[logo, text_block]]

            # 创建表格并设置样式
            table1 = Table(table_data, colWidths=[8 * inch, None])  # 调整列宽度
            table1.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 顶部对齐
                ('TOPPADDING', (0, 0), (-1, -1), 1 * mm),  # 顶部内边距
                ('ALIGN', (1, 0), (1, 0), 'RIGHT'),  # 右侧文字右对齐
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0 * mm),
                #('LEFTPADDING', (0, 0), (0, 0), 0),  # 取消左侧边距
                #('RIGHTPADDING', (1, 0), (1, 0), 0),  # 取消右侧边距
            ]))

            # Add some space below the title
            story.append(Spacer(1, 12))

            # Convert dataframe to list of lists
            data = [merged_df.columns.tolist()] + merged_df.values.tolist()

            # Define margins and usable width
            left_margin = 0.3 * inch
            right_margin = 0.3 * inch
            usable_width = landscape(letter)[0] - left_margin - right_margin  # Total width minus margins

            # Calculate dynamic column widths based on content length
            col_widths = []
            for col in merged_df.columns:
                max_len = max(merged_df[col].astype(str).apply(len).max(), len(str(col)))
                len_num = max(0.7 * inch, min(1.5 * inch, max_len * 0.10 * inch))
                # if len_num > 1.5 * inch:
                #     len_num = 1.5 * inch
                col_widths.append(len_num)
            print(col_widths)
            # Check if col_widths is empty
            if not col_widths:
                print("Column widths are empty.")
                return None  # Early return if column widths are not determined

            total_col_width = 0
            max_cols_per_page = 0
            # max11 = []
            # for width in col_widths:
            #     total_col_width += width
            #     if total_col_width <= usable_width:
            #         max_cols_per_page += 1
            #     else:
            #         total_col_width -= width
            #         max11.append(max_cols_per_page)
            #         print(total_col_width)
            #         total_col_width = width
            #         max_cols_per_page = 1
            # max11.append(max_cols_per_page)
            # print(usable_width)
            # print(max11)
            print("----------------------------")
            print(usable_width)

            # Calculate max columns per page based on usable width
            total_col_width = 0
            max_cols_per_page = 0
            max1 = []
            p = 0
            print(col_widths[0])

            for width in col_widths:
                total_col_width += width
                if total_col_width <= usable_width:
                    max_cols_per_page += 1
                    p = p + 1
                else:
                    max1.append(max_cols_per_page)
                    print(p)
                    insert_column(merged_df, p)
                    p = p + 4
                    total_col_width = width + col_widths[0] + col_widths[1] + col_widths[2]
                    max_cols_per_page = 4
            max1.append(max_cols_per_page)

            print("dddddddddddddddddddddddddddddddd")
            print(max1)

            data = [merged_df.columns.tolist()] + merged_df.values.tolist()

            col_widths = []
            for col in merged_df.columns:
                print("mmmmmmmmmmmmmmmmm")
                print(col)
                max_len = max(merged_df[col].astype(str).apply(len).max(), len(str(col)))
                # print("mmmmmmmmmmmmmmmmm")
                print(max_len)
                #len_num = max(1.0 * inch, min(2.0 * inch, max_len * 0.10 * inch))
                len_num = max(0.7 * inch, min(1.5 * inch, max_len * 0.10 * inch))
                # if len_num > 1.5 * inch:
                #     len_num = 1.5 * inch
                col_widths.append(len_num)

            print(col_widths)

            # Check if col_widths is empty
            if not col_widths:
                print("Column widths are empty.")
                return None  # Early return if column widths are not determined

            # Ensure we have columns to display
            if max_cols_per_page == 0:
                print("No columns fit on the page.")
                return None  # Early return if no columns can be displayed

            # Define row handling logic as before
            max_rows_per_page = 15
            rows = len(data)
            cols = len(data[0])
            num_row_pages = (rows // max_rows_per_page) + 1

            num_col_pages = len(max1)

            start_row = 0

            for row_page in range(num_row_pages):
                # start_row = row_page * max_rows_per_page
                if(start_row == 0):
                    start_row = 1
                else:
                    start_row = start_row + max_rows_per_page -1
                print("Start row", start_row)
                end_row = min(start_row + max_rows_per_page, rows)
                end_col = 0
                for col_page in range(num_col_pages):
                    #start_col = col_page * max1[col_page]
                    start_col = end_col
                    print("----")
                    print(col_page)
                    print(start_col)
                    end_col = end_col + max1[col_page]
                    print(end_col)


                    # Prepare page data and create the table
                    page_data = [row[start_col:end_col] for row in data[start_row+1:end_row]]

                    # 添加 data 的第 0 行的指定列数据
                    header_row = data[1][start_col:end_col]
                    page_data.insert(0, header_row)  # 在 page_data 的开头插入 header_row
                    if start_row!=1:
                        page_data.insert(1, data[2][start_col:end_col])

                    if not page_data or not page_data[0]:  # Check if page_data is empty
                        print("Page data is empty.")
                        continue
                    story.append(table1)
                    titles = Paragraph(title, styles['Title'])
                    story.append(titles)

                    titles_2 =Paragraph(title_2, styles['Title'])
                    story.append(titles_2)

                    style_n = styles["BodyText"]
                    wrapped_data = []
                    for row in page_data:   #换行
                        wrapped_row = []
                        for item in row:
                            paragraph = Paragraph(str(item), style_n)
                            wrapped_row.append(paragraph)
                        wrapped_data.append(wrapped_row)
                    table = Table(wrapped_data, colWidths=col_widths[start_col:end_col], rowHeights=25)
                    print("Table created.")
                    print("================================")
                    print(max1[col_page])
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                        #('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        #('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('TOPPADDING', (0, 0), (-1, -1), 1 * mm),  # 顶部内边距
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 0 * mm),
                        ('SPAN', (0, 0), (0, 1)),
                        ('SPAN', (1, 0), (1, 1)),
                        ('SPAN', (2, 0), (2, 1)),
                        ('SPAN', (3, 0), (max1[col_page]-1, 0)),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ]))
                    story.append(table)

                    # Add page break if not the last row page
                    if row_page < num_row_pages - 1 or col_page < num_col_pages - 1:
                        story.append(PageBreak())


            # Create a spacer to position the footer content
            story.append(Spacer(1, 24))

            #custom_style = ParagraphStyle(name='CustomStyle', fontSize=12, wordWrap='CJK', alignment=1)
            left_content = [
                Paragraph("Confirmed by (BM/HoD):"),
                Spacer(0,12),
                Paragraph("Signature: ________________"),
                Paragraph("Date: October 27, 2024"),
            ]

            # 用户输入安全处理
            #user_note = sanitize_input(user_note)

            # Split note1 by the hyphen and strip whitespace
            note_lines = [line.strip() for line in note1.split('-')]

            # Convert lines into Paragraph objects, adding a hyphen at the beginning of each new line (except the first one)
            note_paragraphs = []
            for i, line in enumerate(note_lines):
                if line:  # Avoid adding empty lines
                    if i > 0:  # Add hyphen for lines after the first
                        note_paragraphs.append(Paragraph("- " + line))
                    else:
                        note_paragraphs.append(Paragraph(line))


            # 右侧内容
            right_content = [
                Paragraph("Note: This is the right section."),
                Paragraph("Name: ______________"),
                #Paragraph(note1),# 添加空行
                *note_paragraphs,
                #Paragraph(user_note, custom_style),  # 用户输入的内容
            ]

            # 将左侧和右侧内容合并成一个新的表格
            combined_data = [[
                Table([[line] for line in left_content], colWidths=[3 * inch]),
                Table([[line] for line in right_content], colWidths=[3 * inch]),
            ]]

            combined_table = Table(combined_data, colWidths=[3.5 * inch, 3.5 * inch])
            combined_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]))

            story.append(combined_table)

            # Add page number
            doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
        except Exception as e:
            print(f"Error generating PDF: {e}")

    return merged_file_path


@app.before_request
def restrict_url():
    # 获取请求的路径
    path = request.path
    # 检查路径是否为 '/'
    if path != '/' and not path.startswith('/static/') or request.query_string:
        # 如果不是，则返回 404 错误
        abort(404)

@app.before_request
def set_server_headers():
    WSGIRequestHandler.server_version = ''
    WSGIRequestHandler.sys_version = ''


@app.after_request
def set_security_headers(response):

    response.headers["X-Frame-Options"] = "DENY"
    response.headers.pop('Server', None)

    # Prevent MIME-sniffing
    response.headers['X-Content-Type-Options'] = 'nosniff'

    return response

def remove_server_header(response: Response):
   response.headers["Server"] = "MyCustomServer"  # Or set to empty string to hide completely
   return response


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':

        title = request.form.get('title', 'merged_data')  # Default title if none provided
        title_2 = request.form.get('title2', 'merged_data')  # Default title if none provided
        note1 = request.form.get('note1', 'merged_data')  # Default title if none provided

        # 检查上传的文件
        if 'users_file' in request.files and 'details_file' in request.files:
            users_file = request.files['users_file']
            details_files = request.files.getlist('details_file')

            if users_file and allowed_file(users_file.filename):
                users_file_name = secure_filename(users_file.filename)
                users_file_path = os.path.join(app.config['UPLOAD_FOLDER'], users_file_name)
                users_file.save(users_file_path)

                details_file_paths = []
                for details_file in details_files:
                    if allowed_file(details_file.filename):
                        details_file_name = secure_filename(details_file.filename)
                        details_file_path = os.path.join(app.config['UPLOAD_FOLDER'], details_file_name)
                        details_file.save(details_file_path)
                        details_file_paths.append(details_file_path)

                #users_file_name = get_timestamped_filename(users_file.filename)
                #details_file_name = get_timestamped_filename(details_file.filename)

                merged_df = merge_excel(users_file_path, details_file_paths)
                if merged_df is None:
                    flash("Error merging xlsm files.", "error")
                    return redirect(url_for('index'))

                if 'download_csv' in request.form:
                    buffer = BytesIO()
                    merged_df.to_csv(buffer, index=False)
                    buffer.seek(0)
                    return send_file(buffer, as_attachment=True, download_name='merged_data.csv', mimetype='text/csv')

                elif 'download_excel' in request.form:
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        merged_df.to_excel(writer, index=False, sheet_name='MergedData')
                    buffer.seek(0)
                    return send_file(buffer, as_attachment=True, download_name='merged_data.xlsx',
                                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

                elif 'download_pdf' in request.form:
                    merged_file_path = save_merged_file(merged_df, 'pdf', title, title_2, note1)
                    buffer = BytesIO()
                    with open(merged_file_path, 'rb') as f:
                        buffer.write(f.read())
                    buffer.seek(0)
                    return send_file(buffer, as_attachment=True, download_name='merged_data.pdf',
                                     mimetype='application/pdf')

                elif 'show_data' in request.form:
                    return render_template('index.html', tables=[merged_df.to_html(classes='data')],
                                           titles=merged_df.columns.values)


        uploaded_files = os.listdir(app.config['UPLOAD_FOLDER'])
        merged_files = os.listdir(app.config['MERGED_FOLDER'])

        return render_template('success.html', uploaded_files=uploaded_files, merged_files=merged_files)
        #return redirect(url_for('index'))

    uploaded_files = os.listdir(app.config['UPLOAD_FOLDER'])
    merged_files = os.listdir(app.config['MERGED_FOLDER'])

    return render_template('index.html', uploaded_files=uploaded_files, merged_files=merged_files)

@app.route('/download/<folder>/<filename>', methods=['GET'])
def download_file(folder, filename):
    folder_path = app.config.get(f'{folder.upper()}_FOLDER')
    file_path = os.path.join(folder_path, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash(f"The file {filename} does not exist.", "error")
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(port= 5000, debug=False)

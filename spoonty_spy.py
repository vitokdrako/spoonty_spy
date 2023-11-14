import random
import openpyxl

def generate_data(room_range, col1_range, col2_range):
    return {
        room: (round(random.uniform(*col1_range), 1), round(random.uniform(*col2_range), 1))
        for room in room_range
        if not str(room).endswith("13") and room != 603
    }

def create_excel_report():
    wb = openpyxl.Workbook()
    ws = wb.active
    room_ranges = [
        list(range(1, 24)),
        list(range(101, 126)),
        list(range(201, 226)),
        list(range(401, 416)),
        list(range(501, 518)),
        list(range(601, 618))
    ]

    col1_range = (11.0, 12.0)
    col2_range = (13.2, 14.0)

    header = ['Kamer', 'kraan', 'douche'] * (len(room_ranges) // 3)
    ws.append(header)

    max_rows = max(len(r) for r in room_ranges)

    for row in range(max_rows):
        row_data = []
        for range_index, room_range in enumerate(room_ranges):
            if row < len(room_range):
                room = room_range[row]
                if str(room).endswith("13") or room == 603:
                    row_data.extend([''] * 3)
                else:
                    data = generate_data([room], col1_range, col2_range)[room]
                    row_data.extend([room] + list(data))
            else:
                row_data.extend([''] * 3)
        ws.append(row_data)

    wb.save('Temp_staat_hotel_2023.xlsx')

create_excel_report()

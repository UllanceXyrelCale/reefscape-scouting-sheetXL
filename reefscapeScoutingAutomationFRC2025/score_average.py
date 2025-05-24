from safety_variable_assign import safety_divider

def leave_avg(wb, reference_sheet, average_sheet):
    raw_game_piece_sheet = wb[reference_sheet]
    average_game_piece_sheet = wb[average_sheet]
    raw_data_sheet = wb['TeleOp - Raw Data']

    for row in range(2,raw_game_piece_sheet.max_row + 1):
        # Variables
        raw_leave_cell = raw_game_piece_sheet.cell(row, 8)
        avg_leave_cell = average_game_piece_sheet.cell(row, 8)
        games = raw_data_sheet.cell(row, 11)

        # Calculation
        avg_leave_cell_value = safety_divider(raw_leave_cell.value, games.value)

        avg_leave_cell.value = avg_leave_cell_value


def game_piece_avg(wb, reference_sheet, average_sheet):
    total_game_piece_sheet = wb[reference_sheet]
    average_game_piece_sheet = wb[average_sheet]

    for row in range(2, total_game_piece_sheet.max_row + 1):
        # Variables
        gp_level_1_cell = total_game_piece_sheet.cell(row, 2) # Corals
        gp_level_2_cell = total_game_piece_sheet.cell(row, 3)
        gp_level_3_cell = total_game_piece_sheet.cell(row, 4)
        gp_level_4_cell = total_game_piece_sheet.cell(row, 5)

        gp_processor_cell = total_game_piece_sheet.cell(row, 6) # Algae
        gp_barge_cell = total_game_piece_sheet.cell(row, 7)

        avg_level_1_cell = average_game_piece_sheet.cell(row, 2) # Corals
        avg_level_2_cell = average_game_piece_sheet.cell(row, 3)
        avg_level_3_cell = average_game_piece_sheet.cell(row, 4)
        avg_level_4_cell = average_game_piece_sheet.cell(row, 5)

        avg_processor_cell = average_game_piece_sheet.cell(row, 6) # Algae
        avg_barge_cell = average_game_piece_sheet.cell(row, 7)

        games = total_game_piece_sheet.cell(row, 11)

        # Calculation
        avg_level_1_value = safety_divider(gp_level_1_cell.value, games.value)
        avg_level_2_value = safety_divider(gp_level_2_cell.value, games.value)
        avg_level_3_value = safety_divider(gp_level_3_cell.value, games.value)
        avg_level_4_value = safety_divider(gp_level_4_cell.value, games.value)
        avg_processor_value = safety_divider(gp_processor_cell.value, games.value)
        avg_barge_value = safety_divider(gp_barge_cell.value, games.value)

        avg_level_1_cell.value = avg_level_1_value
        avg_level_2_cell.value = avg_level_2_value
        avg_level_3_cell.value = avg_level_3_value
        avg_level_4_cell.value = avg_level_4_value
        avg_processor_cell.value = avg_processor_value
        avg_barge_cell.value = avg_barge_value

def total_points_avg(wb, reference_sheet, average_sheet):
    total_points_sheet = wb[reference_sheet]
    average_points_sheet = wb[average_sheet]
    raw_data_sheet = wb['TeleOp - Raw Data']

    for row in range(2, total_points_sheet.max_row + 1):
        # Variables
        total_points_coral_cell = total_points_sheet.cell(row, 9)
        total_points_algae_cell = total_points_sheet.cell(row, 10)
        total_points_teleop_cell = total_points_sheet.cell(row, 11)

        games = raw_data_sheet.cell(row, 11)

        average_points_coral_cell = average_points_sheet.cell(row, 11)
        average_points_algae_cell = average_points_sheet.cell(row, 12)
        average_points_teleop_cell = average_points_sheet.cell(row, 13)

        # Calculation
        average_points_coral_value = safety_divider(total_points_coral_cell.value, games.value)
        average_points_algae_value = safety_divider(total_points_algae_cell.value, games.value)
        average_points_teleop_value = safety_divider(total_points_teleop_cell.value, games.value)

        average_points_coral_cell.value = average_points_coral_value
        average_points_algae_cell.value = average_points_algae_value
        average_points_teleop_cell.value = average_points_teleop_value

def barge_avg(wb, reference_sheet, average_sheet):
    raw_climb_sheet = wb[reference_sheet]
    average_points_sheet = wb[average_sheet]
    raw_data_sheet = wb['TeleOp - Raw Data']

    for row in range(2, raw_climb_sheet.max_row + 1):
        # Variables
        raw_park_cell = raw_climb_sheet.cell(row, 8)
        raw_shallowclimb_cell = raw_climb_sheet.cell(row, 9)
        raw_deepclimb_cell = raw_climb_sheet.cell(row, 10)

        avg_park_cell = average_points_sheet.cell(row, 8)
        avg_shallowclimb_cell = average_points_sheet.cell(row, 9)
        avg_deepclimb_cell = average_points_sheet.cell(row, 10)

        games = raw_data_sheet.cell(row, 11)

        # Calculations
        avg_park_value = safety_divider(raw_park_cell.value, games.value)
        avg_shallowclimb_value = safety_divider(raw_shallowclimb_cell.value, games.value)
        avg_deepclimb_value = safety_divider(raw_deepclimb_cell.value, games.value)

        avg_park_cell.value = avg_park_value
        avg_shallowclimb_cell.value = avg_shallowclimb_value
        avg_deepclimb_cell.value = avg_deepclimb_value
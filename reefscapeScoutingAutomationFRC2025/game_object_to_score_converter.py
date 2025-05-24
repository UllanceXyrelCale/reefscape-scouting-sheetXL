from safety_variable_assign import safety_multiplier, safety_adder


def coral_point_converter(wb, raw_sheet, total_sheet):
    teleOp_rawData = wb[raw_sheet]
    teleOp_totalPoints = wb[total_sheet]

    for row in range(2, teleOp_rawData.max_row + 1):
        # Variables
        raw_level_1_cell = teleOp_rawData.cell(row, 2)
        raw_level_2_cell = teleOp_rawData.cell(row, 3)
        raw_level_3_cell = teleOp_rawData.cell(row, 4)
        raw_level_4_cell = teleOp_rawData.cell(row, 5)

        points_level_1_cell = teleOp_totalPoints.cell(row, 2)
        points_level_2_cell = teleOp_totalPoints.cell(row, 3)
        points_level_3_cell = teleOp_totalPoints.cell(row, 4)
        points_level_4_cell = teleOp_totalPoints.cell(row, 5)
        total_points_coral_cell = teleOp_totalPoints.cell(row, 9)

        # Points Converter
        points_level_1_value = safety_multiplier(raw_level_1_cell.value,2)
        points_level_2_value = safety_multiplier(raw_level_2_cell.value,3)
        points_level_3_value = safety_multiplier(raw_level_3_cell.value,4)
        points_level_4_value = safety_multiplier(raw_level_4_cell.value,5)

        total_points_coral_value = (points_level_1_value + points_level_2_value
                                    + points_level_3_value + points_level_4_value)

        # Assigning values to respective cells
        points_level_1_cell.value = points_level_1_value
        points_level_2_cell.value = points_level_2_value
        points_level_3_cell.value = points_level_3_value
        points_level_4_cell.value = points_level_4_value
        total_points_coral_cell.value = total_points_coral_value

def coral_point_converter_auto(wb, raw_sheet, total_sheet):
    teleOp_rawData = wb[raw_sheet]
    teleOp_totalPoints = wb[total_sheet]

    for row in range(2, teleOp_rawData.max_row + 1):
        # Variables
        raw_level_1_cell = teleOp_rawData.cell(row, 2)
        raw_level_2_cell = teleOp_rawData.cell(row, 3)
        raw_level_3_cell = teleOp_rawData.cell(row, 4)
        raw_level_4_cell = teleOp_rawData.cell(row, 5)
        raw_leave_cell = teleOp_rawData.cell(row, 8)

        points_level_1_cell = teleOp_totalPoints.cell(row, 2)
        points_level_2_cell = teleOp_totalPoints.cell(row, 3)
        points_level_3_cell = teleOp_totalPoints.cell(row, 4)
        points_level_4_cell = teleOp_totalPoints.cell(row, 5)
        total_points_coral_cell = teleOp_totalPoints.cell(row, 9)
        points_leave_cell = teleOp_totalPoints.cell(row, 8)

        # Points Converter
        points_level_1_value = safety_multiplier(raw_level_1_cell.value,3)
        points_level_2_value = safety_multiplier(raw_level_2_cell.value,4)
        points_level_3_value = safety_multiplier(raw_level_3_cell.value,6)
        points_level_4_value = safety_multiplier(raw_level_4_cell.value,7)
        points_leave_cell_value = safety_multiplier(raw_leave_cell.value, 3)

        total_points_coral_value = (points_level_1_value + points_level_2_value
                                    + points_level_3_value + points_level_4_value)

        # Assigning values to respective cells
        points_level_1_cell.value = points_level_1_value
        points_level_2_cell.value = points_level_2_value
        points_level_3_cell.value = points_level_3_value
        points_level_4_cell.value = points_level_4_value
        points_leave_cell.value = points_leave_cell_value
        total_points_coral_cell.value = total_points_coral_value

def algae_point_converter(wb, raw_sheet, total_sheet):
    teleOp_rawData = wb[raw_sheet]
    teleOp_totalPoints = wb[total_sheet]

    for row in range(2, teleOp_rawData.max_row + 1):
        # Variables
        raw_processor_cell = teleOp_rawData.cell(row, 6)
        raw_bargenet_cell = teleOp_rawData.cell(row, 7)

        points_processor_cell = teleOp_totalPoints.cell(row, 6)
        points_bargenet_cell = teleOp_totalPoints.cell(row, 7)
        total_points_algae_cell = teleOp_totalPoints.cell(row, 10)

        # Point Converter
        points_processor_value = safety_multiplier(raw_processor_cell.value,6)
        points_bargenet_value = safety_multiplier(raw_bargenet_cell.value,4)

        total_points_algae_value = (points_processor_value + points_bargenet_value)

        # Assigning values to respective cells
        points_processor_cell.value = points_processor_value
        points_bargenet_cell.value = points_bargenet_value
        total_points_algae_cell.value = total_points_algae_value

def algae_point_converter(wb, raw_sheet, total_sheet):
    teleOp_rawData = wb[raw_sheet]
    teleOp_totalPoints = wb[total_sheet]

    for row in range(2, teleOp_rawData.max_row + 1):
        # Variables
        raw_processor_cell = teleOp_rawData.cell(row, 6)
        raw_bargenet_cell = teleOp_rawData.cell(row, 7)

        points_processor_cell = teleOp_totalPoints.cell(row, 6)
        points_bargenet_cell = teleOp_totalPoints.cell(row, 7)
        total_points_algae_cell = teleOp_totalPoints.cell(row, 10)

        # Point Converter
        points_processor_value = safety_multiplier(raw_processor_cell.value,6)
        points_bargenet_value = safety_multiplier(raw_bargenet_cell.value,4)

        total_points_algae_value = (points_processor_value + points_bargenet_value)

        # Assigning values to respective cells
        points_processor_cell.value = points_processor_value
        points_bargenet_cell.value = points_bargenet_value
        total_points_algae_cell.value = total_points_algae_value

def barge_point_converter(wb, raw_sheet, total_sheet):
    teleOp_rawData = wb[raw_sheet]
    teleOp_totalPoints = wb[total_sheet]

    for row in range(2, teleOp_rawData.max_row + 1):
        # Variables
        raw_park_cell = teleOp_rawData.cell(row, 8)
        raw_shallowclimb_cell = teleOp_rawData.cell(row, 9)
        raw_deepclimb_cell = teleOp_rawData.cell(row, 10)

        total_points_barge_cell = teleOp_totalPoints.cell(row, 8)

        # Calculation
        points_park_value = safety_multiplier(raw_park_cell.value, 2)
        points_shallowclimb_value = safety_multiplier(raw_shallowclimb_cell.value, 6)
        points_deepclimb_value = safety_multiplier(raw_deepclimb_cell.value, 12)

        total_points_barge_cell.value = safety_adder(points_park_value, points_shallowclimb_value, points_deepclimb_value, 0, 0)

def total_point_converter(wb, total_sheet):
    teleOp_totalPoints = wb[total_sheet]

    for row in range(2, teleOp_totalPoints.max_row + 1):
        # Variables
        total_coral_points_cell = teleOp_totalPoints.cell(row, 9)
        total_algae_points_cell = teleOp_totalPoints.cell(row, 10)
        total_barge_points_cell = teleOp_totalPoints.cell(row, 8)
        total_teleop_points_cell = teleOp_totalPoints.cell(row, 11)

        try:
            total_teleop_points_cell.value = total_coral_points_cell.value + total_algae_points_cell.value + total_barge_points_cell.value
        except TypeError:
            total_teleop_points_cell.value = total_coral_points_cell.value + total_algae_points_cell.value

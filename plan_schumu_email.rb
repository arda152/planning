# Arda Arman
# September 2020

require "csv"
require "write_xlsx"
require "mail"
require "io/console"

# TIMESLOT DATABASE

# Time slots for the whole week, manually written to be flexible with different opening hours, they contain:
# => duration : the duration of the practice slot, in hours
# => slot_name: the name of the day and the time slot
# => available_rooms: all of the room names that are available in this time slot (empty in the beginning)
# => slot_score: capacity - requests to practice in the room. Used later in the algorithm.
timeslots = [
    {"timeslot_index" => 0, "slot_score" => 0, "duration" => 2, "slot_name" => "Montag 08:00-10:00", "available_rooms" => []},
    {"timeslot_index" => 1, "slot_score" => 0, "duration" => 2, "slot_name" => "Montag 10:00-12:00", "available_rooms" => []},
    {"timeslot_index" => 2, "slot_score" => 0, "duration" => 2, "slot_name" => "Montag 12:00-14:00", "available_rooms" => []},
    {"timeslot_index" => 3, "slot_score" => 0, "duration" => 2, "slot_name" => "Montag 14:00-16:00", "available_rooms" => []},
    {"timeslot_index" => 4, "slot_score" => 0, "duration" => 2, "slot_name" => "Montag 16:00-18:00", "available_rooms" => []},
    {"timeslot_index" => 5, "slot_score" => 0, "duration" => 2, "slot_name" => "Montag 18:00-20:00", "available_rooms" => []},
    {"timeslot_index" => 6, "slot_score" => 0, "duration" => 2, "slot_name" => "Montag 20:00-22:00", "available_rooms" => []},
    {"timeslot_index" => 7, "slot_score" => 0, "duration" => 2, "slot_name" => "Dienstag 08:00-10:00", "available_rooms" => []},
    {"timeslot_index" => 8, "slot_score" => 0, "duration" => 2, "slot_name" => "Dienstag 10:00-12:00", "available_rooms" => []},
    {"timeslot_index" => 9, "slot_score" => 0, "duration" => 2, "slot_name" => "Dienstag 12:00-14:00", "available_rooms" => []},
    {"timeslot_index" => 10, "slot_score" => 0, "duration" => 2, "slot_name" => "Dienstag 14:00-16:00", "available_rooms" => []},
    {"timeslot_index" => 11, "slot_score" => 0, "duration" => 2, "slot_name" => "Dienstag 16:00-18:00", "available_rooms" => []},
    {"timeslot_index" => 12, "slot_score" => 0, "duration" => 2, "slot_name" => "Dienstag 18:00-20:00", "available_rooms" => []},
    {"timeslot_index" => 13, "slot_score" => 0, "duration" => 2, "slot_name" => "Dienstag 20:00-22:00", "available_rooms" => []},
    {"timeslot_index" => 14, "slot_score" => 0, "duration" => 2, "slot_name" => "Mittwoch 08:00-10:00", "available_rooms" => []},
    {"timeslot_index" => 15, "slot_score" => 0, "duration" => 2, "slot_name" => "Mittwoch 10:00-12:00", "available_rooms" => []},
    {"timeslot_index" => 16, "slot_score" => 0, "duration" => 2, "slot_name" => "Mittwoch 12:00-14:00", "available_rooms" => []},
    {"timeslot_index" => 17, "slot_score" => 0, "duration" => 2, "slot_name" => "Mittwoch 14:00-16:00", "available_rooms" => []},
    {"timeslot_index" => 18, "slot_score" => 0, "duration" => 2, "slot_name" => "Mittwoch 16:00-18:00", "available_rooms" => []},
    {"timeslot_index" => 19, "slot_score" => 0, "duration" => 2, "slot_name" => "Mittwoch 18:00-20:00", "available_rooms" => []},
    {"timeslot_index" => 20, "slot_score" => 0, "duration" => 2, "slot_name" => "Mittwoch 20:00-22:00", "available_rooms" => []},
    {"timeslot_index" => 21, "slot_score" => 0, "duration" => 2, "slot_name" => "Donnerstag 08:00-10:00", "available_rooms" => []},
    {"timeslot_index" => 22, "slot_score" => 0, "duration" => 2, "slot_name" => "Donnerstag 10:00-12:00", "available_rooms" => []},
    {"timeslot_index" => 23, "slot_score" => 0, "duration" => 2, "slot_name" => "Donnerstag 12:00-14:00", "available_rooms" => []},
    {"timeslot_index" => 24, "slot_score" => 0, "duration" => 2, "slot_name" => "Donnerstag 14:00-16:00", "available_rooms" => []},
    {"timeslot_index" => 25, "slot_score" => 0, "duration" => 2, "slot_name" => "Donnerstag 16:00-18:00", "available_rooms" => []},
    {"timeslot_index" => 26, "slot_score" => 0, "duration" => 2, "slot_name" => "Donnerstag 18:00-20:00", "available_rooms" => []},
    {"timeslot_index" => 27, "slot_score" => 0, "duration" => 2, "slot_name" => "Donnerstag 20:00-22:00", "available_rooms" => []},
    {"timeslot_index" => 28, "slot_score" => 0, "duration" => 2, "slot_name" => "Freitag 08:00-10:00", "available_rooms" => []},
    {"timeslot_index" => 29, "slot_score" => 0, "duration" => 2, "slot_name" => "Freitag 10:00-12:00", "available_rooms" => []},
    {"timeslot_index" => 30, "slot_score" => 0, "duration" => 2, "slot_name" => "Freitag 12:00-14:00", "available_rooms" => []},
    {"timeslot_index" => 31, "slot_score" => 0, "duration" => 2, "slot_name" => "Freitag 14:00-16:00", "available_rooms" => []},
    {"timeslot_index" => 32, "slot_score" => 0, "duration" => 2, "slot_name" => "Freitag 16:00-18:00", "available_rooms" => []},
    {"timeslot_index" => 33, "slot_score" => 0, "duration" => 2, "slot_name" => "Freitag 18:00-20:00", "available_rooms" => []},
    {"timeslot_index" => 34, "slot_score" => 0, "duration" => 2, "slot_name" => "Freitag 20:00-22:00", "available_rooms" => []},
    {"timeslot_index" => 35, "slot_score" => 0, "duration" => 3, "slot_name" => "Samstag 09:00-12:00", "available_rooms" => []},
    {"timeslot_index" => 36, "slot_score" => 0, "duration" => 2, "slot_name" => "Samstag 12:00-14:00", "available_rooms" => []},
    {"timeslot_index" => 37, "slot_score" => 0, "duration" => 2, "slot_name" => "Samstag 14:00-16:00", "available_rooms" => []},
    {"timeslot_index" => 38, "slot_score" => 0, "duration" => 2, "slot_name" => "Samstag 16:00-18:00", "available_rooms" => []},
    {"timeslot_index" => 39, "slot_score" => 0, "duration" => 2, "slot_name" => "Samstag 18:00-20:00", "available_rooms" => []},
    {"timeslot_index" => 40, "slot_score" => 0, "duration" => 2, "slot_name" => "Samstag 20:00-22:00", "available_rooms" => []},
    {"timeslot_index" => 41, "slot_score" => 0, "duration" => 2, "slot_name" => "Sonntag 12:00-14:00", "available_rooms" => []},
    {"timeslot_index" => 42, "slot_score" => 0, "duration" => 2, "slot_name" => "Sonntag 14:00-16:00", "available_rooms" => []},
    {"timeslot_index" => 43, "slot_score" => 0, "duration" => 2, "slot_name" => "Sonntag 16:00-18:00", "available_rooms" => []},
    {"timeslot_index" => 44, "slot_score" => 0, "duration" => 2, "slot_name" => "Sonntag 18:00-20:00", "available_rooms" => []},
    {"timeslot_index" => 45, "slot_score" => 0, "duration" => 2, "slot_name" => "Sonntag 20:00-22:00", "available_rooms" => []}
]

# INPUT FROM 2 CSV DOCUMENTS (TEACHERS & STUDENTS)

# FIRST TABLE, FROM THE STUDENTS

# Read the student response csv and create the students array
# The students array contains hashes for each student with the values:
# => name: full name of the student
# => email: email address of the student
# => free_time: a list of student's free time, the values are strings, they must be in the exact same format as those in timeslots["slot_name"] field
# => practice_time: a list of student's practice slots, this is a combination of room name and time slot strings
# => total_duration: the number of hours in the complete week for each student

csvstudents = CSV.read("studentresponse.csv")
students = []

# First row of csv is not needed
student_index = 1
while student_index < csvstudents.length
    # Hash for the new student
    new_student_name = csvstudents[student_index][2] + " " + csvstudents[student_index][3]
    new_student_email = csvstudents[student_index][1]
    new_student = {"name" => new_student_name, "email" => new_student_email, "free_time" => [], "practice_time" => [], "total_duration" => 0, "free_time_in_hours" => 0}

    # Ask student if they are free on timeslot
    timeslots.each_with_index do |slot, slot_index|
        # CSV data is shifted by 4 cells to the right to match the time timeslots
        if csvstudents[student_index][slot_index + 4] != "Ich kann nicht üben."
            new_student["free_time"] << slot["slot_name"]
            # Free time of student in hours
            new_student["free_time_in_hours"] += slot["duration"]
            # Decrease slot_score by one, somebody wants to practice there
            slot["slot_score"] -= 1
        end
    end
    students << new_student
    student_index += 1
end

# SECOND TABLE, FROM THE TEACHERS

# Create the list of rooms and mark the free times in the timeslots array
# Rooms list will keep track of each rooms own plan independently from other parts of the algorithm
# This way we can create separate plans for each room.
rooms = []
csvteacher = CSV.read("teacherresponse.csv")
# First row of csv is not needed
room_index = 1
while room_index < csvteacher.length
    new_room_name = csvteacher[room_index][2] + " " + csvteacher[room_index][3]
    new_room = {"name" => new_room_name, "reserved_practice" => []}
    timeslots.each_with_index do |slot, slot_index|
        # CSV data is shifted by 5 cells to the right to match the time timeslots
        if csvteacher[room_index][slot_index + 5] != "Belegt"
            # Add the available room to the available rooms section
            slot["available_rooms"] << new_room_name
            # Increase the slot score by one, we have additional capacity now
            slot["slot_score"] += 1
        end
    end
    # Add the new rooms name to the rooms Hash
    rooms << new_room
    room_index += 1
end


# Calculate the total time available in the whole week for practice.
total_time_available = 0
timeslots.each do |timeslot|
    total_time_available += timeslot["duration"] * timeslot["available_rooms"].length
end




# PLANNING ALGORITHM

# The program goes through all time slots
# Each room is given to the free student who has the least total reserved hours until that moment
# This makes sure everything stays equal

# The only condition: the algorithm needs to have rooms with lots of competition (negative slot score) at the end
# These rooms with lots of competition are used to equal out students with less practice time
# Since the slot score is very low, we know that many students can come at these time slots, so equalizing becomes easy
# In the same way, using the rooms with positive slot score (lots of free space, no competition) will be helpful
# Giving priority to these zero competition rooms makes sure that we actually find the people who wanted to practice there
# This also makes sure that zero competition rooms are given as soon as possible, otherwise they can remain too empty


# Order the timeslots with decreasing score. Rooms that have lots of free space will be given first
timeslots.sort! do |x, y|
    y["slot_score"] <=> x["slot_score"]
end

# For each timeslot in the list of free rooms
timeslots.each do |timeslot|
    # For every student who is free in this slot
    freestudents = students.select do |student|

        # Schulmusik students are allowed to practice 2 hours per day
        day_of_possible_timeslot = timeslot["slot_name"].split(" ")[0]

        # None of previously reserved slots have the same day name with the current possible slot
        no_reservations_on_same_day = student["practice_time"].none? do |reserved_timeslot|
            reserved_timeslot[0]["slot_name"].split(" ")[0] == day_of_possible_timeslot
        end

        # Student has free time and no reservation on the same day
        has_time = (student["free_time"].include? timeslot["slot_name"])
        (has_time && no_reservations_on_same_day)
    end

    freestudents.each do |free_student|
        # If there are still rooms available in the timeslot
        if timeslot["available_rooms"].length > 0
            # Place the student in the first room available
            room_name = timeslot["available_rooms"][0]
            free_student["practice_time"] << [timeslot, room_name]
            # Update student's total time
            free_student["total_duration"] += timeslot["duration"]
            # Update room data
            room_name = timeslot["available_rooms"][0]

            selected_room = rooms.select do |room|
                room["name"] == room_name
            end
            selected_room[0]["reserved_practice"] << [timeslot, free_student["name"]]

            # Remove the room from the timeslot
            timeslot["available_rooms"].delete_at(0)
        end
    end

    # After each timeslot, reorganize the students, so the one with the least practice gets the next room
    students.sort! do |x, y|
        x["total_duration"] <=> y["total_duration"]
    end
end

# CONSOLE REPORT ABOUT LOST HOURS AND TOTAL PRACTICE DURATIONS

# Calculate the total time reserved for practice for later calculations
total_time_reserved = 0

puts "STUDIERENDE"
students.each do |student|
    student_wish_in_hours = student["free_time_in_hours"]
    puts student["name"] + " hat " + student["total_duration"].to_s + " Stunden Übezeit. (" + (100 * (student["total_duration"].to_f / student_wish_in_hours)).to_i.to_s + "%)"
    total_time_reserved += student["total_duration"]
end
puts ""

puts "LEERE ZIMMER"
timeslots.select {|timeslot| timeslot["available_rooms"].length != 0}.each do |slot|
    puts slot["slot_name"] + "(slot_score:" + slot["slot_score"].to_s + "):"
    puts slot["available_rooms"]
    puts ""
end

# Calculate the lost hours
lost_hours = total_time_available - total_time_reserved
puts total_time_reserved.to_s + " Stunden insgesamt reserviert zum Üben, " + lost_hours.to_s + " Stunden (" + (100 * (lost_hours.to_f / total_time_available)).to_i.to_s + "%) bleiben unbenutzt."
puts ""

# What is the approximate limit for students who get less than 100% of their selection?
approximate_limit = total_time_available

students.each do |student|
    total_time_reserved += student["total_duration"]
    student_wish_in_hours = student["free_time_in_hours"]
    if ((student["total_duration"] < student_wish_in_hours))
        approximate_limit = student["total_duration"] < approximate_limit ? student["total_duration"] : approximate_limit
    end
end
puts "Jeder Studierende, der weniger als 100% seiner Wahl bekommen hat, hat zumindest " + approximate_limit.to_s + " Stunden Übezeit."

# OUTPUT AS CSV DOCUMENT

# Sort the results using timeslot index
# This makes sure that the weekly plan is sorted Montag -> Sonntag

rooms.each do |room|
    room["reserved_practice"].sort! do |x, y|
        x[0]["timeslot_index"] <=> y[0]["timeslot_index"]
    end
end

students.each do |student|
    student["practice_time"].sort! do |x, y|
        x[0]["timeslot_index"] <=> y[0]["timeslot_index"]
    end
end

# Ask the date for the tables
puts "Datum:"
week_date = gets.chomp.to_s


# Output rooms
# Create a new Excel workbook
filename = "./Zimmerplan.xlsx"
workbook = WriteXLSX.new(filename)
row_index = 1
worksheet = workbook.add_worksheet

rooms.each do |room|

    # Add a worksheet
    worksheet.fit_to_pages(1, 1)

    format_bold = workbook.add_format
    format_bold.set_bold()
    format_bold.set_size(20)

    format_normal = workbook.add_format
    format_normal.set_size(14)

    worksheet.set_column(0, 1, 30)
    worksheet.write(row_index, 0, room["name"], format_bold)
    worksheet.write(row_index, 1, week_date, format_bold)
    # Write a formatted and unformatted string, row and column notation.
    row_index += 1
    room["reserved_practice"].each do |slot|
        worksheet.write(row_index, 0, slot[0]["slot_name"], format_normal)
        worksheet.write(row_index, 1, slot[1], format_normal)
        row_index += 1
    end
    row_index += 1
end
workbook.close



# Output students
Dir.mkdir("./output_stud")
students_csv_array = []
students.each do |student|
    # Create a new Excel workbook
    filename = "./output_stud/" + student["name"] + ".xlsx"
    workbook = WriteXLSX.new(filename)

    # Add a worksheet
    worksheet = workbook.add_worksheet
    worksheet.fit_to_pages(1, 1)

    # Formatiing
    format_bold = workbook.add_format
    format_bold.set_bold()
    format_bold.set_size(20)

    format_normal = workbook.add_format
    format_normal.set_size(14)

    # Student name and date
    worksheet.set_column(0, 1, 30)
    worksheet.write(0, 0, student["name"], format_bold)

    worksheet.write(0, 1, week_date, format_bold)
    row_index = 1

    # Practice slots

    student["practice_time"].each do |slot|
        room_name = slot[1]
        worksheet.write(row_index, 0, slot[0]["slot_name"], format_normal)
        worksheet.write(row_index, 1, slot[1], format_normal)
        row_index += 1
    end
    workbook.close
end


# SEND THE MAILS TO THE STUDENTS

puts "\n" + students.count.to_s + " Pläne sind bereit."
puts "Bitte prüfen und dann Eingabetaste drücken, um die E-Mails zu senden."
enter = gets
puts "raumplanung.hmt@gmail.com Passwort:"
password = STDIN.noecho(&:gets).chomp

options = {
    :address              => "smtp.gmail.com",
    :port                 => 587,
    :domain               => 'gmail.com',
    :user_name            => 'raumplanung.hmt',
    :password             => password,
    :authentication       => 'plain',
    :enable_starttls_auto => true
}

Mail.defaults do
  delivery_method :smtp, options
end


student_count = 1
students.each do |student|
    puts "Email an: " + student["name"] + "  " + "[" + student_count.to_s + "/" + students.length.to_s + "]"
    filepath = "./output_stud/" + student["name"] + ".xlsx"

    mail = Mail.new do
        from "Johannes Keller"
        to student["email"]
        subject "Übezimmerverteilung"
        body "Hallo " + student["name"].split(" ")[0] + "," + "\n" + "im Anhang findest du deinen Wochenplan.\nViele Grüße\nJohannes Keller"
        add_file filepath
    end

    mail.charset = "UTF-8"

    mail.deliver!
    student_count += 1
end

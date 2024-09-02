#!/usr/bin/env python
# coding: utf-8

import numpy as np
import pandas as pd

import random
from copy import deepcopy

import streamlit as st
import io

import warnings
warnings.filterwarnings("ignore")


# In[3]:


# Data input
def read_input(file):
    # Membaca berbagai sheet dari file Excel
    df_sesi = pd.read_excel(file, sheet_name='Sesi')
    df_ruang = pd.read_excel(file, sheet_name='Ruangan')
    df_makul = pd.read_excel(file, sheet_name='Makul')
    df_tidak_ujian = pd.read_excel(file, sheet_name='Tidak Ujian')
    df_butuh_lab = pd.read_excel(file, sheet_name='Butuh Lab')
    return df_sesi, df_ruang, df_makul, df_tidak_ujian, df_butuh_lab

# Fungsi utama Streamlit
def main():
    st.title("Aplikasi Penjadwalan Ujian")
    st.title("Departemen Statistika Bisnis ITS")

    # Mengunggah file Excel
    uploaded_file = st.file_uploader("Unggah file Excel", type=["xlsx"])
    
    if uploaded_file is not None:
        # Baca file Excel yang diunggah
        df_sesi, df_ruang, df_makul, df_tidak_ujian, df_butuh_lab = read_input(uploaded_file)

        df_makul.fillna(method='ffill', inplace = True)
        df_sesi.fillna(method='ffill', inplace = True)
        
        df_tidak_ujian.columns = ['Butuh Lab']
        merged = df_butuh_lab.merge(df_tidak_ujian, on=['Butuh Lab'], how='outer', indicator=True)
        df_butuh_lab = merged[merged['_merge'] == 'left_only'][df_butuh_lab.columns].reset_index(drop=True)
        df_tidak_ujian.columns = ['Tidak Ujian']
        df_makul = df_makul[~df_makul['Mata Kuliah'].isin(df_tidak_ujian['Tidak Ujian'])].reset_index(drop=True)
        
        df_sesi['wkwk'] = 1
        df_ruang['wkwk'] = 1
        
        slot_jadwal = pd.merge(df_sesi, df_ruang, on='wkwk')
        del df_sesi['wkwk']
        del df_ruang['wkwk']
        del slot_jadwal['wkwk']
        
        df_makul['Semester'] = df_makul['Semester'].astype(int)
        
        # Genetic Algorithm Parameters
        population_size = 50
        generations = 100
        mutation_rate = 0.1
        crossover_rate = 0.7
        
        def generate_initial_population(df_sesi, df_makul):
            population = []
            
            for _ in range(population_size):
                individual = []
                # Dictionary to keep track of the count of subjects scheduled in each time slot
                time_slot_counts = {(hari, jam): 0 for hari, jam in zip(df_sesi['Hari'], df_sesi['Jam'])}
        
                for matakuliah, group in df_makul.groupby('Mata Kuliah'):
                    # Find a valid session that does not exceed the constraint of 2 subjects per time slot
                    valid_sesi = [sesi for sesi in df_sesi.index 
                                  if time_slot_counts[(df_sesi.loc[sesi, 'Hari'], df_sesi.loc[sesi, 'Jam'])] < 2]
                    
                    # Randomly choose a valid session
                    sesi = random.choice(valid_sesi)
                    individual.append((df_sesi.loc[sesi, 'Hari'], df_sesi.loc[sesi, 'Jam'], matakuliah, group['Semester'].values[0]))
                    
                    # Update the count for this time slot
                    time_slot_counts[(df_sesi.loc[sesi, 'Hari'], df_sesi.loc[sesi, 'Jam'])] += 1
                
                population.append(individual)
            
            return population
        
        def fitness_function(individual, df_sesi, df_makul, df_butuh_lab):
            fitness = 0
            semester_distribution = {semester: 0 for semester in df_makul['Semester'].unique()}
            lab_constraints = {day: 0 for day in df_sesi['Hari'].unique()}
            
            # Check constraints
            for entry in individual:
                hari, jam, matakuliah, semester = entry
                
                # Update semester distribution
                semester_distribution[semester] += 1
                
                # Check if the course needs a lab
                if matakuliah in df_butuh_lab['Butuh Lab'].values:
                    lab_constraints[hari] += 1
            
            # Fitness evaluation
            # Penalize if a semester has more than 2 hours of exams on the same day
            for semester, count in semester_distribution.items():
                if count > 2:
                    fitness -= (count - 2) * 10
            
            # Penalize if lab constraints are violated
            for hari, count in lab_constraints.items():
                if count > df_sesi[df_sesi['Hari'] == hari].shape[0]:
                    fitness -= (count - df_sesi[df_sesi['Hari'] == hari].shape[0]) * 10
        
            # Reward for balancing semester distribution across days
            fitness += len(semester_distribution.keys()) - len(set(semester_distribution.values()))
            
            return fitness
        
        def crossover(parent1, parent2, df_sesi):
            while True:  # Loop until we find a valid crossover
                crossover_point = random.randint(0, len(parent1) - 1)
                child1 = parent1[:crossover_point] + parent2[crossover_point:]
                child2 = parent2[:crossover_point] + parent1[crossover_point:]
                
                # Cek validitas anak
                if is_valid_schedule(child1, df_sesi) and is_valid_schedule(child2, df_sesi):
                    return child1, child2
        
        def is_valid_schedule(individual, df_sesi):
            # Cek apakah jadwal valid (tidak ada lebih dari 2 mata kuliah di satu slot waktu)
            time_slot_counts = {(hari, jam): 0 for hari, jam in zip(df_sesi['Hari'], df_sesi['Jam'])}
            
            for entry in individual:
                hari, jam = entry[0], entry[1]
                time_slot_counts[(hari, jam)] += 1
                
                if time_slot_counts[(hari, jam)] > 2:
                    return False
            
            return True
        
        def mutate(individual, df_sesi):
            if random.random() < mutation_rate:
                mutate_idx = random.randint(0, len(individual) - 1)
                hari, jam, matakuliah, semester = individual[mutate_idx]
                
                # Cek sesi yang valid dengan jumlah maksimal 2 mata kuliah
                time_slot_counts = {(hari, jam): 0 for hari, jam in zip(df_sesi['Hari'], df_sesi['Jam'])}
                
                # Hitung jumlah mata kuliah di setiap slot waktu dalam individu saat ini
                for entry in individual:
                    time_slot_counts[(entry[0], entry[1])] += 1
                
                # Cari sesi baru yang tidak melanggar constraint
                valid_sesi = [sesi for sesi in df_sesi.index 
                              if time_slot_counts[(df_sesi.loc[sesi, 'Hari'], df_sesi.loc[sesi, 'Jam'])] < 2]
                
                # Jika ada sesi yang valid, lakukan mutasi
                if valid_sesi:
                    new_sesi = random.choice(valid_sesi)
                    individual[mutate_idx] = (df_sesi.loc[new_sesi, 'Hari'], df_sesi.loc[new_sesi, 'Jam'], matakuliah, semester)
        
        def genetic_algorithm(df_sesi, df_makul, df_butuh_lab):
            population = generate_initial_population(df_sesi, df_makul)
            best_individual = None
            best_fitness = float('-inf')
            
            for generation in range(generations):
                population_fitness = [(individual, fitness_function(individual, df_sesi, df_makul, df_butuh_lab)) for individual in population]
                population_fitness.sort(key=lambda x: x[1], reverse=True)
                
                # Keep track of the best individual
                if population_fitness[0][1] > best_fitness:
                    best_individual = population_fitness[0][0]
                    best_fitness = population_fitness[0][1]
                
                # Selection (Tournament Selection)
                new_population = []
                for _ in range(population_size // 2):
                    parent1 = random.choice(population_fitness)[0]
                    parent2 = random.choice(population_fitness)[0]
                    if random.random() < crossover_rate:
                        child1, child2 = crossover(parent1, parent2, df_sesi)
                    else:
                        child1, child2 = parent1, parent2
                    
                    mutate(child1, df_sesi)
                    mutate(child2, df_sesi)
                    
                    new_population.extend([child1, child2])
                
                population = new_population
            
            return best_individual
        
        # Run the genetic algorithm
        best_schedule = genetic_algorithm(df_sesi, df_makul, df_butuh_lab)
        
        jadwal = pd.DataFrame(best_schedule)
        jadwal.columns = ['Hari', 'Jam', 'Mata Kuliah', 'Semester']
        jadwal.sort_values(by=['Hari', 'Jam', 'Mata Kuliah', 'Semester'], inplace=True)  
        
        jadwal_dan_makul = pd.merge(jadwal, df_makul)
        del jadwal_dan_makul['Dosen Pengampu']
        del jadwal_dan_makul['Semester']
        
        df_ruang_lab = df_ruang[df_ruang['Lab'] == 'y'].sort_values(by='Kapasitas', ascending=False).copy()
        df_ruang_non_lab = df_ruang[df_ruang['Lab'] != 'y'].sort_values(by='Kapasitas', ascending=False).copy()
        
        # STEP 2: Define fitness function to minimize the number of rooms with merged classes
        def fitness(schedule, df_butuh_lab):
            fitness_score = 0
            
            for _, row in schedule.iterrows():
                mata_kuliah = row['Mata Kuliah']
                jumlah_mahasiswa = row['Jumlah Mahasiswa']
                kapasitas_ruang = row['Kapasitas']
                is_lab_course = mata_kuliah in df_butuh_lab['Butuh Lab'].values
                
                # Selisih antara kapasitas ruangan dan jumlah mahasiswa
                difference = kapasitas_ruang - jumlah_mahasiswa
                
                # Penilaian untuk mata kuliah yang butuh lab
                if is_lab_course and row['Ruang'] in df_ruang_lab['Ruang'].values:
                    # Jika mata kuliah membutuhkan lab dan ditempatkan di ruang lab
                    fitness_score += (kapasitas_ruang - difference)  # Bonus untuk penggunaan ruang lab
                elif is_lab_course and row['Ruang'] not in df_ruang_lab['Ruang'].values:
                    # Penalti jika mata kuliah yang membutuhkan lab tidak ditempatkan di ruang lab
                    fitness_score -= 10
                
                # Optimalisasi kapasitas: semakin kecil selisihnya, semakin baik
                fitness_score -= difference  # Penalti semakin besar selisih kapasitas dan jumlah mahasiswa
            
            return fitness_score
        
        # Assign students to rooms based on capacity and lab requirements, prioritizing lab-required courses
        def generate_schedule(df_makul, df_ruang, df_butuh_lab):
            schedule = []
            
            # Separate lab and non-lab rooms
            df_ruang_lab = df_ruang[df_ruang['Lab'] == 'y'].sort_values(by='Kapasitas', ascending=False).copy()
            df_ruang_non_lab = df_ruang[df_ruang['Lab'] != 'y'].sort_values(by='Kapasitas', ascending=False).copy()
        
            # Group classes by Mata Kuliah and Hari-Waktu to combine if possible
            grouped_classes = df_makul.groupby(['Hari', 'Jam', 'Mata Kuliah'])
        
            for (hari, jam, mata_kuliah), group in grouped_classes:
                total_students = group['Jumlah Mahasiswa'].sum()
                lab_needed = mata_kuliah in df_butuh_lab['Butuh Lab'].values
                remaining_students = total_students
        
                # Prioritize lab rooms if the course needs it
                if lab_needed:
                    for _, room_row in df_ruang_lab.iterrows():
                        if remaining_students <= 0:
                            break
                        
                        # Check for conflicts in the same room at the same time
                        conflicting = [s for s in schedule if s[0] == hari and s[1] == jam and s[4] == room_row['Ruang']]
                        if conflicting:
                            continue  # Skip if there's a conflict
        
                        assigned_room = room_row['Ruang']
                        room_capacity = room_row['Kapasitas']
        
                        students_assigned = min(remaining_students, room_capacity)
                        remaining_students -= students_assigned
        
                        if students_assigned > 0:
                            schedule.append([hari, jam, mata_kuliah, 'Gabungan', assigned_room, students_assigned, room_capacity])
        
                    # If lab rooms aren't enough, assign to non-lab rooms
                    if remaining_students > 0:
                        for _, room_row in df_ruang_non_lab.iterrows():
                            if remaining_students <= 0:
                                break
                            
                            # Check for conflicts in the same room at the same time
                            conflicting = [s for s in schedule if s[0] == hari and s[1] == jam and s[4] == room_row['Ruang']]
                            if conflicting:
                                continue  # Skip if there's a conflict
        
                            assigned_room = room_row['Ruang']
                            room_capacity = room_row['Kapasitas']
        
                            students_assigned = min(remaining_students, room_capacity)
                            remaining_students -= students_assigned
        
                            if students_assigned > 0:
                                schedule.append([hari, jam, mata_kuliah, 'Gabungan', assigned_room, students_assigned, room_capacity])
        
                # For non-lab courses, assign to non-lab rooms first
                else:
                    for _, room_row in df_ruang_non_lab.iterrows():
                        if remaining_students <= 0:
                            break
                        
                        # Check for conflicts in the same room at the same time
                        conflicting = [s for s in schedule if s[0] == hari and s[1] == jam and s[4] == room_row['Ruang']]
                        if conflicting:
                            continue  # Skip if there's a conflict
        
                        assigned_room = room_row['Ruang']
                        room_capacity = room_row['Kapasitas']
        
                        students_assigned = min(remaining_students, room_capacity)
                        remaining_students -= students_assigned
        
                        if students_assigned > 0:
                            schedule.append([hari, jam, mata_kuliah, 'Gabungan', assigned_room, students_assigned, room_capacity])
        
                    # If non-lab rooms aren't enough, assign to lab rooms as a last resort
                    if remaining_students > 0:
                        for _, room_row in df_ruang_lab.iterrows():
                            if remaining_students <= 0:
                                break
                            
                            # Check for conflicts in the same room at the same time
                            conflicting = [s for s in schedule if s[0] == hari and s[1] == jam and s[4] == room_row['Ruang']]
                            if conflicting:
                                continue  # Skip if there's a conflict
        
                            assigned_room = room_row['Ruang']
                            room_capacity = room_row['Kapasitas']
        
                            students_assigned = min(remaining_students, room_capacity)
                            remaining_students -= students_assigned
        
                            if students_assigned > 0:
                                schedule.append([hari, jam, mata_kuliah, 'Gabungan', assigned_room, students_assigned, room_capacity])
        
            return pd.DataFrame(schedule, columns=['Hari', 'Waktu', 'Mata Kuliah', 'Kelas', 'Ruang', 'Jumlah Mahasiswa', 'Kapasitas'])
        
        # Genetic algorithm loop to optimize the schedule
        best_schedule = None
        best_fitness = float('-inf')
        
        # Generate fixed schedules and then optimize them
        fixed_schedule = generate_schedule(jadwal_dan_makul, df_ruang, df_butuh_lab)
        
        for generation in range(5):
            schedule = generate_schedule(jadwal_dan_makul, df_ruang, df_butuh_lab)
            
            current_fitness = fitness(schedule, df_butuh_lab)
            
            if current_fitness > best_fitness:
                best_fitness = current_fitness
                best_schedule = schedule
        
        jadwal_makul_dan_ruang = pd.DataFrame(best_schedule).sort_values(by = ['Hari', 'Waktu', 'Mata Kuliah'])
        del jadwal_makul_dan_ruang['Kelas']
        
        df_jadwal_ruang = jadwal_makul_dan_ruang[['Hari', 'Waktu', 'Mata Kuliah', 'Ruang']]
        df_jadwal_ruang['Dosen Pengampu'] = df_jadwal_ruang['Mata Kuliah'].map(df_makul[['Mata Kuliah', 'Dosen Pengampu']].drop_duplicates().set_index('Mata Kuliah')['Dosen Pengampu'])
        
        # Mengambil semua dosen unik dari daftar pengampu
        dosen_all = list(set(', '.join(df_jadwal_ruang['Dosen Pengampu']).split(', ')))
        
        # Fungsi untuk membuat jadwal pengawas secara acak
        def generate_schedule(df_jadwal, df_ruang):
            schedule = df_jadwal.copy()
            # Struktur data untuk melacak pengawas yang sudah ditugaskan pada setiap hari dan jam tertentu
            assigned_supervisors = {hari: {waktu: [] for waktu in df_jadwal['Waktu'].unique()} for hari in df_jadwal['Hari'].unique()}
            
            # Assign pengawas
            schedule['Pengawas'] = schedule.apply(lambda row: assign_pengawas(row, df_ruang, assigned_supervisors), axis=1)
            return schedule
        
        def assign_pengawas(row, df_ruang, assigned_supervisors):
            ruang = row['Ruang']
            hari = row['Hari']
            waktu = row['Waktu']
            jumlah_pengawas = df_ruang[df_ruang['Ruang'] == ruang]['Jumlah Pengawas'].values[0]
            dosen_pengampu = row['Dosen Pengampu'].split(', ')
            
            # Mencari pengawas yang merupakan pengampu dan belum ditugaskan pada hari dan waktu tersebut
            possible_supervisors = [d for d in dosen_all if d in dosen_pengampu and d not in assigned_supervisors[hari][waktu]]
            
            # Jika tidak ada pengawas yang merupakan pengampu dan memenuhi syarat, pilih dari dosen lain yang belum ditugaskan
            if len(possible_supervisors) < jumlah_pengawas:
                possible_supervisors += [d for d in dosen_all if d not in dosen_pengampu and d not in assigned_supervisors[hari][waktu]]
            
            # Mengacak dan memilih pengawas sesuai dengan jumlah yang diperlukan
            selected_supervisors = random.sample(possible_supervisors, jumlah_pengawas)
            
            # Update pengawas yang telah ditugaskan untuk hari dan waktu tersebut
            assigned_supervisors[hari][waktu].extend(selected_supervisors)
            
            return ', '.join(selected_supervisors)
        
        # Fungsi untuk menghitung fitness
        def calculate_fitness(schedule):
            dosen_supervise_count = {dosen: 0 for dosen in dosen_all}
            
            # Hitung jumlah mengawas setiap dosen
            for pengawas in schedule['Pengawas']:
                for p in pengawas.split(', '):
                    dosen_supervise_count[p] += 1
            
            # Hitung fitness sebagai 1 / (1 + std)
            supervise_counts = list(dosen_supervise_count.values())
            fitness = 1 / (1 + np.std(supervise_counts))
            return fitness
        
        # Fungsi untuk melakukan seleksi, crossover, dan mutasi
        def genetic_algorithm(df_jadwal, df_ruang, generations=100, population_size=100):
            population = [generate_schedule(df_jadwal, df_ruang) for _ in range(population_size)]
            best_schedule = None
            best_fitness = 0
            
            for generation in range(generations):
                # Hitung fitness dari seluruh populasi
                fitness_scores = [calculate_fitness(schedule) for schedule in population]
                
                # Simpan individu dengan fitness terbaik
                max_fitness_idx = np.argmax(fitness_scores)
                if fitness_scores[max_fitness_idx] > best_fitness:
                    best_fitness = fitness_scores[max_fitness_idx]
                    best_schedule = deepcopy(population[max_fitness_idx])
                
                # Seleksi individu berdasarkan fitness
                selected_population = random.choices(population, weights=fitness_scores, k=population_size // 2)
                
                # Crossover (pertukaran pengawas antar dua jadwal)
                new_population = []
                for i in range(0, len(selected_population), 2):
                    if i+1 < len(selected_population):
                        parent1, parent2 = selected_population[i], selected_population[i+1]
                        child1, child2 = crossover(parent1, parent2)
                        new_population.extend([child1, child2])
                    else:
                        new_population.append(selected_population[i])
                
                # Mutasi (mengubah pengawas secara acak)
                population = [mutate(schedule, df_ruang) for schedule in new_population]
            
            return best_schedule, best_fitness
        
        def crossover(parent1, parent2):
            crossover_point = random.randint(0, len(parent1))
            child1 = deepcopy(parent1)
            child2 = deepcopy(parent2)
            for i in range(crossover_point, len(parent1)):
                child1['Pengawas'].iloc[i] = parent2['Pengawas'].iloc[i]
                child2['Pengawas'].iloc[i] = parent1['Pengawas'].iloc[i]
            return child1, child2
        
        def mutate(schedule, df_ruang):
            mutated_schedule = deepcopy(schedule)
            mutate_idx = random.randint(0, len(schedule) - 1)
            hari = mutated_schedule.iloc[mutate_idx]['Hari']
            waktu = mutated_schedule.iloc[mutate_idx]['Waktu']
            assigned_supervisors = {hari: {waktu: [] for waktu in df_jadwal_ruang['Waktu'].unique()} for hari in df_jadwal_ruang['Hari'].unique()}
            mutated_schedule['Pengawas'].iloc[mutate_idx] = assign_pengawas(mutated_schedule.iloc[mutate_idx], df_ruang, assigned_supervisors)
            return mutated_schedule
        
        # Menjalankan algoritma genetika
        best_schedule, best_fitness = genetic_algorithm(df_jadwal_ruang, df_ruang)
        
        # Fungsi untuk menghitung jumlah mengawas setiap pengawas pada setiap hari
        def count_supervisions_per_day(schedule):
            dosen_supervise_count_per_day = {dosen: {hari: 0 for hari in schedule['Hari'].unique()} for dosen in dosen_all}
        
            # Iterasi setiap baris jadwal untuk menghitung jumlah pengawasan per hari
            for index, row in schedule.iterrows():
                hari = row['Hari']
                pengawas_list = row['Pengawas'].split(', ')
                for pengawas in pengawas_list:
                    if pengawas in dosen_supervise_count_per_day:
                        dosen_supervise_count_per_day[pengawas][hari] += 1
        
            # Mengubah hasil ke DataFrame untuk tampilan yang lebih baik
            df_count_per_day = pd.DataFrame(dosen_supervise_count_per_day).transpose()
            return df_count_per_day
        
        # Menghitung dan menampilkan jumlah mengawas per hari untuk setiap dosen
        supervision_counts = count_supervisions_per_day(best_schedule)
        supervision_counts['Total'] = supervision_counts.sum(axis=1)
        
        jadwal_lengkap = pd.DataFrame(best_schedule)
        jadwal_lengkap['Jumlah Mahasiswa'] = jadwal_makul_dan_ruang['Jumlah Mahasiswa']
        
        jadwal_lengkap = jadwal_lengkap[[
                        'Hari',
                        'Waktu',
                        'Mata Kuliah',
                        # 'Kelas',
                        'Dosen Pengampu',
                        'Pengawas',
                        'Ruang',
                        'Jumlah Mahasiswa',
                        # 'Kapasitas',
                        ]]
        
        def alokasi_mahasiswa(jadwal_lengkap, df_makul):
            # Salin jadwal_lengkap agar tidak mengubah data asli
            jadwal_updated = jadwal_lengkap.copy()
        
            # Buat dictionary untuk menyimpan jumlah mahasiswa yang belum diassign per mata kuliah dan kelas
            sisa_mahasiswa = {}
            kelas_semester = {}  # Dictionary tambahan untuk menyimpan informasi semester setiap kelas
        
            for mata_kuliah, group in df_makul.groupby('Mata Kuliah'):
                sisa_mahasiswa[mata_kuliah] = group[['Kelas', 'Jumlah Mahasiswa']].set_index('Kelas').to_dict()['Jumlah Mahasiswa']
                # Tambahkan informasi semester untuk setiap kelas
                kelas_semester[mata_kuliah] = group[['Kelas', 'Semester']].set_index('Kelas').to_dict()['Semester']
        
            # Loop untuk mengalokasikan mahasiswa ke ruangan
            for idx, row in jadwal_updated.iterrows():
                mata_kuliah = row['Mata Kuliah']
                ruangan_mahasiswa = row['Jumlah Mahasiswa']
        
                # Mendapatkan kelas yang terkait dengan mata kuliah ini
                if mata_kuliah not in sisa_mahasiswa:
                    continue
                
                kelas_info = sisa_mahasiswa[mata_kuliah]
                semester_info = kelas_semester[mata_kuliah]  # Ambil informasi semester terkait kelas
                kelas_assigned = []
        
                # Cek apakah ada jumlah mahasiswa yang sama dengan kapasitas ruangan
                found_exact_match = False
                for kelas, jumlah_mahasiswa in list(kelas_info.items()):
                    if jumlah_mahasiswa == ruangan_mahasiswa and jumlah_mahasiswa > 0:
                        # Jika jumlah mahasiswa kelas sama dengan kapasitas ruangan, alokasikan langsung
                        semester = semester_info[kelas]  # Ambil semester kelas
                        kelas_assigned.append(f"{semester}{kelas} ({jumlah_mahasiswa})")
                        kelas_info[kelas] -= jumlah_mahasiswa
                        ruangan_mahasiswa = 0  # Ruang sudah terisi penuh
                        found_exact_match = True
                        break  # Keluar dari loop karena ruang sudah penuh
        
                # Jika ditemukan kecocokan jumlah mahasiswa dengan kapasitas ruangan, lewati alokasi berikutnya
                if not found_exact_match:
                    # Alokasi mahasiswa ke ruangan dengan strategi terbaik yang memungkinkan
                    kelas_sorted = sorted(kelas_info.items(), key=lambda x: -x[1])  # Sortir berdasarkan jumlah mahasiswa terbesar
                    for kelas, jumlah_mahasiswa in kelas_sorted:
                        if ruangan_mahasiswa == 0:
                            break
        
                        if jumlah_mahasiswa > 0:
                            semester = semester_info[kelas]  # Ambil semester kelas
                            if jumlah_mahasiswa <= ruangan_mahasiswa:
                                # Jika seluruh kelas bisa muat dalam ruang, masukkan seluruh kelas
                                kelas_assigned.append(f"{semester}{kelas} ({jumlah_mahasiswa})")
                                ruangan_mahasiswa -= jumlah_mahasiswa
                                kelas_info[kelas] -= jumlah_mahasiswa
                            else:
                                # Jika sebagian kelas bisa muat dalam ruang, masukkan sebagian kelas
                                kelas_assigned.append(f"{semester}{kelas} ({ruangan_mahasiswa})")
                                kelas_info[kelas] -= ruangan_mahasiswa
                                ruangan_mahasiswa = 0
        
                # Simpan kelas yang telah dialokasikan pada jadwal
                jadwal_updated.at[idx, 'Kelas'] = ", ".join(kelas_assigned)
            
            return jadwal_updated
        
        # Menjalankan fungsi alokasi
        jadwal_terupdate = alokasi_mahasiswa(jadwal_lengkap, df_makul)
        
        a = pd.DataFrame(jadwal_terupdate)
        a['Kapasitas'] = jadwal_makul_dan_ruang['Kapasitas']
        
        slot_jadwal = slot_jadwal.rename(columns={'Jam': 'Waktu'})
        
        b = pd.merge(slot_jadwal, a, how='left', on=['Hari', 'Waktu', 'Ruang', 'Kapasitas'])
        b = b.rename(columns={'Dosen Pengampu': 'Dosen'})
        b['Lab'].replace('y', 'lab', inplace=True)
        b = b[[
                        'Hari',
                        'Waktu',
                        'Mata Kuliah',
                        'Kelas',
                        'Dosen',
                        'Pengawas',
                        'Ruang',
                        'Lab',
                        'Jumlah Mahasiswa',
                        'Kapasitas',
                        ]]

        # Menyimpan output ke dalam Excel dan memungkinkan pengunduhan
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            b.to_excel(writer, sheet_name='Jadwal', index=False)
            supervision_counts.to_excel(writer, sheet_name='Rekap Jaga')
        
        # Buat tombol untuk mengunduh file Excel hasilnya
        st.write("Silahkan Unduh Excel Berikut:")
        st.download_button(
            label="Unduh Jadwal Terupdate",
            data=output.getvalue(),
            file_name="jadwal_ujian.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Proses data dan tampilkan hasil
        st.write("Data Sesi:", df_sesi)
        st.write("Data Ruangan:", df_ruang)
        st.write("Data Mata Kuliah:", df_makul)
        st.write("Data Tidak Ujian:", df_tidak_ujian)
        st.write("Data Butuh Lab:", df_butuh_lab)

        # Tampilkan hasil
        st.write("Jadwal Terupdate dengan Alokasi Mahasiswa:")
        st.dataframe(a)
        st.write("Rekap Jaga:")
        st.dataframe(supervision_counts)

if __name__ == "__main__":
    main()

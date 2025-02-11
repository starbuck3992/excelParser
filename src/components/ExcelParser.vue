<template>
  <div class="bg-white rounded-lg shadow-md p-4 m-auto w-96">
    <div class="block flex flex-col w-fit mx-auto">
      <label class="p-4 rounded-lg text-black text-center bg-gray-200 hover:bg-gray-300 cursor-pointer">
        <span>Upload Excel</span>
        <input ref="fileInput" type="file" @input="handleFileUpload" accept=".xlsx, .xls" />
      </label>
      <div v-show="filteredDates">
        <h1>
          Filtered Dates
        </h1>
        {{filteredDates}}
      </div>
      <div class="flex items-center gap-2 ml-auto">
        <button v-show="filteredDates" :class="['p-4 rounded-lg bg-red-500 text-white hover:opacity-500 mt-12']" @click="clear">Reset</button>
        <button v-show="filteredDates" :class="['p-4 rounded-lg bg-indigo-500 text-white hover:opacity-500 mt-12']" @click="downloadExcel">Download Excel</button>
      </div>

    </div>
  </div>
</template>

<script setup lang="ts">
import {onMounted, ref} from 'vue';
import * as XLSX from 'xlsx';

onMounted(async () => {
  await fetchHolidays();
});

const fileInput = ref();

const holidays = ref();

const filteredDates = ref();

const fetchHolidays = async () => {
  const country = "CZ";
  const dateFrom = new Date().toISOString().split('T')[0];
  const dateTo = new Date(new Date().setFullYear(new Date().getFullYear() + 3)).toISOString().split('T')[0];

  const url = `https://openholidaysapi.org/PublicHolidays?countryIsoCode=${country}&validFrom=${dateFrom}&validTo=${dateTo}`;

  try {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`Response status: ${response.status}`);
    }

    const json = await response.json();

    holidays.value = json.map((item)=> {
      return item.startDate;
    });
  } catch (error) {
    console.error('Error fetching holidays:', error);
  }
}

const handleFileUpload = (event) => {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    if (!e.target?.result) return;

    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Read first sheet
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Convert to YYYY-MM-DD format
    const parsedDates = rows.flat().map(formatDate).filter(Boolean);

    // Adjust holidays to the next working day
    filteredDates.value = parsedDates.map(adjustToNextWorkingDay);
  };
  reader.readAsArrayBuffer(file);
};

// Format date from Excel
const formatDate = (date) => {
  if (!date) return null;
  let dt;

  if (typeof date === 'number') {
    dt = XLSX.SSF.parse_date_code(date);
    return `${dt.y}-${String(dt.m).padStart(2, '0')}-${String(dt.d).padStart(2, '0')}`;
  }

  try {
    dt = new Date(date);
    return dt.toISOString().split('T')[0];
  } catch {
    return null;
  }
};

// Adjust holidays to the next working day (Mon-Fri, non-holiday)
const adjustToNextWorkingDay = (dateStr) => {
  let date = new Date(dateStr);

  // If it's not a holiday, return as-is
  if (!holidays.value.includes(dateStr)) return dateStr;

  // Move forward until we hit a working day
  do {
    date.setDate(date.getDate() + 1);
  } while (holidays.value.includes(date.toISOString().split('T')[0]) || isWeekend(date));

  return date.toISOString().split('T')[0];
};


// Check if a date is a weekend (Saturday or Sunday)
const isWeekend = (date) => date.getDay() === 0 || date.getDay() === 6;

// Function to download filtered dates as an Excel file
const downloadExcel = () => {
  const ws = XLSX.utils.aoa_to_sheet([["Filtered Dates"], ...filteredDates.value.map(date => [date])]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Filtered Dates");

  XLSX.writeFile(wb, "filtered_dates.xlsx");
};

const clear = () => {
  fileInput.value.value = '';
  filteredDates.value = null;
};
</script>

<style lang="css">
input[type="file"] {
  display: none;
}

</style>

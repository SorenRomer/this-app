<template>
  <div class="h-screen w-screen flex flex-col bg-slate-100 align-left">
    <div class="m-10">
      <input type="file" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        @change="handleFileUpload" />
    </div>
    <div class="m-5">
      <select v-model="location">
        <option value="">Select location</option>
        <option value="odense">Odense</option>
        <option value="gentofte">Gentofte</option>
        <option value="rksk">Rksk</option>
      </select>
    </div>
    <!-- <div class="m-5" v-if="location === 'odense' || location === 'gentofte'"> -->
    <div class="m-5">
      <label>
        <input type="radio" v-model="language" value="danish" />
        Danish
      </label>
      <label>
        <input type="radio" v-model="language" value="english" />
        English
      </label>
    </div>
    <div class="m-5">
      <button class="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded" :disabled="!file || !language" @click="processFile">Process File</button>
    </div>
  </div>
</template>

<script>
import { read, utils, writeFile } from 'xlsx';

export default {
  data() {
    return {
      location: '',
      language: '',
      showLanguage: false,
      oddOrEven: 0
    }
  },
  methods: {
    handleFileUpload(event) {
      this.file = event.target.files[0];
    },
    processFile() {
      const file = this.file;
      if (!file) {
        alert("Please select a file");
        return;
      }
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = utils.sheet_to_json(sheet);

        const outputData = json.map((row) => {
          return {
            UserName: row.Email.trim(),
            Name: row.Name.replace(/ .*/, '').trim(),
            Surname: row.Name.indexOf(" ") !== -1 ? row.Name.split(' ').slice(1).join(' ') : "Missing last name",
            EmailAddress: row.Email.trim(),
            PhoneNumber: this.cleanPhoneNo(row.PhoneNumber),
            Password: this.location.toLocaleLowerCase(),
            AssignedRoleNames: "Citizen",
            AssignedCaseworker: (this.location.toLocaleLowerCase() === "odense") ? this.getCaseworker() : "",
            AssignedCourseflow: this.parseCourse(row.Course, this.location, this.language),
            CourseflowStartdate: this.getNextMonday(this.language),

            SendActivationMail: "True",
          };
        });

        // Convert the output data to an Excel workbook
        const outputSheet = utils.json_to_sheet(outputData);
        const outputWorkbook = utils.book_new();
        utils.book_append_sheet(outputWorkbook, outputSheet, 'Sheet1');

        // Save the output workbook to a file
        writeFile(outputWorkbook, 'output.xlsx');
      };
      reader.readAsArrayBuffer(file);
    },
    getNextMonday(language) {
      const date = new Date();
      const daysUntilNextMonday = ((7 - date.getDay()) % 7) + 1;
      date.setDate(date.getDate() + daysUntilNextMonday);

      const year = date.getFullYear();
      let month = date.getMonth() + 1;
      let day = date.getDate();

      if (language === 'english') {
        month = month < 10 ? '0' + month : month;
        day = day < 10 ? '0' + day : day;
        return `${month}-${day}-${year}`;
      } else {
        month = month < 10 ? '0' + month : month;
        day = day < 10 ? '0' + day : day;
        return `${day}-${month}-${year}`;
      }
    },
    getCaseworker() {
      this.oddOrEven++;
      return this.oddOrEven % 2 ? 'hmv@ballisager.com' : 'aso@ballisager.com'
    },
    parseCourse(course, location, language) {
      if (location.toLocaleLowerCase() === "odense" ) {
        if (language.toLocaleLowerCase() === "english") {
          return "Odense_ENG_jun22"
        } else {
          if (course && course.indludes('fagl')) {
            return "Odense-ufaglaert_DK_sep22"
          } else {
            return "Odense_DK_jun22"
          }          
        }
      } else if (location.toLocaleLowerCase() === "gentofte") {
        if (language.toLocaleLowerCase() === "english") {
          return "Gentofte_Int_jun22"
        } else {
          if (course.includes('enerel')) {
            return "Gentofte_Gen_sep22"
          } else if (course.includes('specialist')) {
            return ""
          } else if (course.includes('mittend')) {
            return "Gentofte_Dim_sep22"
          }
        }          
      } else if (location.toLocaleLowerCase() === "rksk") {
        return "RKSK_Deltid_sep22"
      }

      return ""
    },
    cleanPhoneNo(phoneNo) {
      let phoneStr = phoneNo + ""
      if (phoneStr.includes("+45")) {
        phoneStr = phoneStr.replace("+45", "")
      }
      return (phoneStr.replace(/\D/g,''))
    }
  },
  watch: {
    selectedLocation(newValue) {
      if (newValue === 'odense' || newValue === 'gentofte') {
        this.showLanguage = true;
      } else {
        this.showLanguage = false;
        this.language = '';
      }
    }
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}
</style>

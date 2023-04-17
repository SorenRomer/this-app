<template>
  <div class="w-screen flex flex-col align-left items-center">
    <div class="container">
      <div class="my-10">
        <h3 class="mb-5 text-lg"><b>Step 1:</b> Klargør din kildefil</h3>
        <p class="mb-5">
          Først skal du udfylde din import-fil. Her skal være følgende fire kolonner:
        </p>
        <div class="flex rounded bg-green-50 border-2 border-green-400 justify-items-stretch items-start mb-5">
          <div class="p-2 border-r-2 border-green-400 w-1/4">
            <b>Name</b>
            <br>
            Jens Jensen
          </div>
          <div class="p-2 border-r-2 border-green-400 w-1/4">
            <b>Email</b>
            <br>
            jens.jensen@jobspot.dk
          </div>
          <div class="p-2 border-r-2 border-green-400 w-1/4">
            <b>PhoneNumer</b>
            <br>
            30639434
          </div>
          <div class="p-2 w-1/4">
            <b>Course</b>
            <br>
            Generel / Øvrige
          </div>
        </div>
        <p class="mb-5">
          Det er ikke et problem, hvis der er flere kolonner. Du kan med andre ord nøjes med, at tilpasse overskrifterne på kolonnerne direkte i de filer du modtager fra jobcenteret.
        </p>
        <p class="mb-5">
          Det vigtige er, at kolonnerne har navnene: <b>Name</b>, <b>Email</b>, <b>PhoneNumber</b> og <b>Course</b>.
        </p>
        <p class="mb-5">
          For Odense kan Course være tomt, med mindre det er en ufaglært. Er det det, vil det stå i kolonnen 'Noter' på Excel-filen fra Odense. I så fald skal der bare stå 'ufaglært' i Course-kolonnen.
        </p>
        <p class="mb-5">
          Gentofte har følgende fire kurser: 'Øvrige/generel', 'Dimittend', 'Leder/specialister' og 'Engelsk'. 'Engelsk' skal håndteres i en separat fil, jf. Step 4. De øvrige kan bare copy-pastes direkte fra mail'en fra Gentofte. Husk dog at tjekke for eventuelle tastefejl. Det er forekommet enkelte gange, at de har skrevet det forkert, og så vil AssignedCourseflow-feltet i output-filen være tomt.
        </p>
        <p class="mb-5">
          Hvis du foretrækker at bruge en skabelon, så kan du downloade en <a class="text-blue-500 hover:text-blue-800 decoration-inherit underline" href="/assets/xlsxtemplate.xlsx" download>her</a>.
        </p>        
      </div>
      <hr class="my-10">
      <div class="my-10">   
        <h3 class="mb-5 text-lg"><b>Step 2:</b> Upload din kildefil</h3>     
        <input 
          type="file" 
          accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"          
          @change="handleFileUpload" 
          class="relative m-0 block min-w-0 flex-auto cursor-pointer rounded border border-solid border-neutral-300 bg-clip-padding px-3 py-[0.32rem] font-normal leading-[2.15] text-neutral-700 transition duration-300 ease-in-out file:-mx-3 file:-my-[0.32rem] file:cursor-pointer file:overflow-hidden file:rounded-none file:border-0 file:border-solid file:border-inherit file:bg-neutral-100 file:px-3 file:py-[0.32rem] file:text-neutral-700 file:transition file:duration-150 file:ease-in-out file:[border-inline-end-width:1px] file:[margin-inline-end:0.75rem] hover:file:bg-neutral-200 focus:border-primary focus:text-neutral-700 focus:shadow-te-primary focus:outline-none dark:border-neutral-600 dark:text-neutral-200 dark:file:bg-neutral-700 dark:file:text-neutral-100 dark:focus:border-primary"
        />
        <span class="text-xs"><i>Filen skal være af formatet xlsx</i></span>
      </div>
      <hr class="my-10">
      <div class="my-10">
        <div class="description mb-5">
          <h3 class="mb-5 text-lg"><b>Step 3:</b> Vælg jobcenter</h3>
          <p>
            Nu skal du vælge hvilket jobcenter du importerer for. Dette afgør hvilke module der bliver output'et.
          </p>
        </div>
        <select v-model="location" class="rounded bg-white border-neutral-300 border py-2 pl-4 pr-6 text-neutral-600">
          <option value="">Vælg jobcenter</option>
          <option value="odense">Odense</option>
          <option value="gentofte">Gentofte</option>
          <option value="rksk">Rksk</option>
        </select>
      </div>
      <hr class="my-10">
      <!-- <div class="m-5" v-if="location === 'odense' || location === 'gentofte'"> -->
      <div class="my-10">
        <div class="description mb-5">
          <h3 class="mb-5 text-lg"><b>Step 3:</b> Vælg sprog</h3>
          <p class="mb-5">
            Nu skal du vælge hvilket sprog du importerer til. Dette sikrer, at du får den rette datoformatering
          </p>
          <p class="mb-5 p-2 pb-4 bg-sky-100 border-2 border-sky-500 rounded"><b class="text-sky-500">OBS FOR GENTOFTE-IMPORT!</b><br>Husk at du skal skifte sprog på selve platformen inden du importerer, så det matcher importfilen.</p>
        </div>
        <div class="mb-5 cursor-pointer">
          <label class="cursor-pointer">
            <input type="radio" v-model="language" value="danish" class="cursor-pointer" />
            <span class="ml-4">Dansk</span>
          </label>
        </div>
        <div>
          <label class="cursor-pointer">
            <input type="radio" v-model="language" value="english" class="cursor-pointer" />
            <span class="ml-4">English</span>
          </label>
        </div>        
      </div>
      <hr class="my-10">
      <div class="mb-20">
        <button  
          class="bg-blue-500 hover:bg-blue-700 text-white font-bold py-3 px-5 text-lg rounded disabled:bg-slate-300 disabled:text-slate-100 uppercase" 
          :disabled="!file || !language"
          @click="processFile"
        >
          Processer Fil
        </button>
      </div>
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
            UserName: { v: row.Email.trim(), t: 's' },
            Name: { v: row.Name.replace(/ .*/, '').trim(), t: 's' },
            Surname: { v: row.Name.indexOf(" ") !== -1 ? row.Name.split(' ').slice(1).join(' ') : "Missing last name", t: 's' },
            EmailAddress: { v: row.Email.trim(), t: 's' },
            PhoneNumber: { v: this.cleanPhoneNo(row.PhoneNumber), t: 's' },
            Password: { v: this.location.toLocaleLowerCase(), t: 's' },
            AssignedRoleNames: { v: "Citizen", t: 's' },
            AssignedCaseworker: { v: (this.location.toLocaleLowerCase() === "odense") ? this.getCaseworker() : "", t: 's' },
            AssignedCourseflow: { v: this.parseCourse(row.Course, this.location, this.language), t: 's' },
            CourseflowStartdate: { v: String(this.getNextMonday(this.language)), t: 's' },
            SendActivationMail: { v: "True", t: 's' }
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
        return " " + `${month}-${day}-${year}`;
      } else {
        month = month < 10 ? '0' + month : month;
        day = day < 10 ? '0' + day : day;
        return " " + `${day}-${month}-${year}`;
      }
    },
    getCaseworker() {
      this.oddOrEven++;
      return this.oddOrEven % 2 ? 'hmv@ballisager.com' : 'aso@ballisager.com'
    },
    parseCourse(course, location, language) {
      if (location.toLocaleLowerCase() === "odense") {
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
      return (phoneStr.replace(/\D/g, ''))
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
body {
  background-color: rgb(241 245 249 / var(--tw-bg-opacity));
}
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}

.container {
  max-width: 800px !important;
  min-width: 640px;
}
</style>

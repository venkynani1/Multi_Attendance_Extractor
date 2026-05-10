<template>
  <div class="attendance-page">
    <div class="background-grid"></div>
    <main class="report-shell">
      <header class="hero-section">
        <div class="hero-copy">
          <span class="hero-badge"><Sparkles :size="15" /> Smart Attendance Extractor</span>
          <h1>Attendance Report Generator</h1>
          <p>Consolidate Teams and Webex CSV attendance, merge duplicate durations, match nominations, identify unmatched records, and export a polished Excel report in one flow.</p>
        </div>
        <aside class="hero-visual" aria-label="Workflow preview">
          <div class="hero-card-row">
            <Upload :size="20" />
            <div>
              <strong>CSV Upload</strong>
              <span>{{ selectedFilesCount }} attendance files</span>
            </div>
          </div>
          <div class="hero-card-row">
            <Users :size="20" />
            <div>
              <strong>Nomination Match</strong>
              <span>{{ nominationOption || 'Awaiting selection' }}</span>
            </div>
          </div>
          <div class="hero-card-row">
            <FileSpreadsheet :size="20" />
            <div>
              <strong>Excel Export</strong>
              <span>{{ exportReady ? 'Ready' : 'Pending extraction' }}</span>
            </div>
          </div>
        </aside>
      </header>

      <section class="stepper-card" aria-label="Report progress">
        <div
          v-for="step in stepItems"
          :key="step.id"
          :class="['step-item', { active: activeStep === step.id, complete: step.id < activeStep }]"
        >
          <span class="step-icon">
            <component :is="step.icon" :size="17" />
          </span>
          <span>{{ step.label }}</span>
        </div>
      </section>

      <div v-if="errorMessage" class="notice notice-error">
        <XCircle :size="18" />
        <span>{{ errorMessage }}</span>
      </div>
      <div v-if="successMessage" class="notice notice-success">
        <CheckCircle :size="18" />
        <span>{{ successMessage }}</span>
      </div>

      <section class="dashboard-card">
        <div class="section-heading">
          <div>
            <h2>Attendance Source</h2>
            <p>Choose the platform for this report. One source stays active for a clean export.</p>
          </div>
          <span v-if="selectedTypeLabel" class="status-pill">{{ selectedTypeLabel }}</span>
        </div>

        <div class="source-grid">
          <article
            :class="['source-card', { selected: attendanceType === 'Teams', disabled: attendanceType === 'Webex' && attendanceFiles.length }]"
            tabindex="0"
            role="button"
            aria-label="Select Microsoft Teams attendance"
            @click="handleAttendanceTypeSelection('Teams')"
            @keydown.enter="handleAttendanceTypeSelection('Teams')"
          >
            <div class="source-icon teams"><Users :size="24" /></div>
            <div>
              <h3>Microsoft Teams</h3>
              <p>Supports non-standard Teams CSVs with participant sections.</p>
              <span class="file-type">CSV attendance reports</span>
            </div>
            <label class="drop-zone" @dragover.prevent @drop.prevent="handleAttendanceDrop($event, 'Teams')">
              <Upload :size="28" />
              <input
                ref="teamsInput"
                type="file"
                accept=".csv,text/csv"
                multiple
                :disabled="attendanceType === 'Webex' && attendanceFiles.length > 0"
                @change="handleAttendanceUpload($event, 'Teams')"
              />
              <strong>Drop CSV files here or browse</strong>
              <span>Drop CSV files here or browse</span>
            </label>
          </article>

          <article
            :class="['source-card', { selected: attendanceType === 'Webex', disabled: attendanceType === 'Teams' && attendanceFiles.length }]"
            tabindex="0"
            role="button"
            aria-label="Select Cisco Webex attendance"
            @click="handleAttendanceTypeSelection('Webex')"
            @keydown.enter="handleAttendanceTypeSelection('Webex')"
          >
            <div class="source-icon webex"><Mail :size="24" /></div>
            <div>
              <h3>Cisco Webex</h3>
              <p>Groups reconnects and guest joins by normalized attendee name.</p>
              <span class="file-type">CSV attendance reports</span>
            </div>
            <label class="drop-zone" @dragover.prevent @drop.prevent="handleAttendanceDrop($event, 'Webex')">
              <Upload :size="28" />
              <input
                ref="webexInput"
                type="file"
                accept=".csv,text/csv"
                multiple
                :disabled="attendanceType === 'Teams' && attendanceFiles.length > 0"
                @change="handleAttendanceUpload($event, 'Webex')"
              />
              <strong>Drop CSV files here or browse</strong>
              <span>Drop CSV files here or browse</span>
            </label>
          </article>
        </div>

        <div v-if="attendanceFiles.length" class="file-panel">
          <div class="file-panel-header">
            <strong>{{ selectedFilesCount }} files selected</strong>
            <span>Multiple sessions will be consolidated into one Excel sheet.</span>
          </div>
          <div class="file-chip-grid">
            <div v-for="(file, index) in attendanceFiles" :key="`${file.name}-${index}`" class="file-chip">
              <FileCheck :size="17" />
              <div>
                <strong>{{ file.name }}</strong>
                <span>{{ formatFileSize(file.size) }}</span>
              </div>
              <button type="button" aria-label="Remove file" @click.stop="removeAttendanceFile(index)">
                <XCircle :size="16" />
              </button>
            </div>
          </div>
        </div>
      </section>

      <section v-if="attendanceFiles.length" class="dashboard-card options-card">
        <div class="section-heading">
          <div>
            <h2>Report Options</h2>
            <p>Choose nomination matching and the minimum duration required for Present.</p>
          </div>
        </div>

        <div class="options-grid">
          <div class="field-group">
            <span class="field-label">Do you have nominations?</span>
            <div class="segmented">
              <button type="button" :class="{ selected: nominationOption === 'Yes' }" @click="handleNominationOptionChange('Yes')">
                <CheckCircle :size="16" />
                Yes
              </button>
              <button type="button" :class="{ selected: nominationOption === 'No' }" @click="handleNominationOptionChange('No')">
                <XCircle :size="16" />
                No
              </button>
            </div>
          </div>

          <div class="field-group">
            <span class="field-label">Minimum Stay</span>
            <div class="duration-pills">
              <button
                v-for="duration in [0, 15, 30, 45, 60]"
                :key="duration"
                type="button"
                :class="{ selected: minDurationStay === duration }"
                @click="minDurationStay = duration"
              >
                <Timer :size="15" />
                {{ duration }} mins
              </button>
            </div>
          </div>

          <div v-if="nominationOption === 'Yes'" class="nomination-upload">
            <span class="field-label">Nomination Excel File</span>
            <label class="excel-zone">
              <FileSpreadsheet :size="24" />
              <input ref="nominationInput" type="file" accept=".xlsx,.xls" @change="handleNominationUpload" />
              <strong>Upload nomination workbook</strong>
              <span>{{ nominationHint }}</span>
            </label>
            <div v-if="nominationFile" class="nomination-chip">
              <FileSpreadsheet :size="17" />
              <div>
                <strong>{{ nominationFile.name }}</strong>
                <span>{{ formatFileSize(nominationFile.size) }}</span>
              </div>
              <button type="button" aria-label="Remove nomination file" @click="removeNominationFile">
                <XCircle :size="16" />
              </button>
            </div>
          </div>
        </div>
      </section>

      <section class="command-card">
        <div class="action-bar">
          <button type="button" class="btn btn-primary" :disabled="!canExtract || loading" @click="prepareFinalAttendance">
            <span v-if="loading" class="spinner"></span>
            <FileCheck v-else :size="18" />
            {{ loading ? 'Extracting...' : 'Extract Attendance' }}
          </button>
          <button type="button" class="btn btn-secondary" :disabled="!exportReady || loading" @click="exportExcel">
            <Download :size="18" />
            Export Excel
          </button>
          <button type="button" class="btn btn-danger-ghost" :disabled="loading" @click="resetForm">
            <RotateCcw :size="18" />
            Reset
          </button>
        </div>
      </section>

      <section v-if="exportReady" class="summary-grid">
        <div class="summary-card" style="--delay: 0ms">
          <Users :size="22" />
          <strong>{{ summary.totalNominatedEmployees }}</strong>
          <span>Total Employees</span>
          <small>Rows in final report</small>
        </div>
        <div class="summary-card" style="--delay: 70ms">
          <FileCheck :size="22" />
          <strong>{{ summary.totalUploadedAttendanceFiles }}</strong>
          <span>Uploaded Sessions</span>
          <small>CSV sessions processed</small>
        </div>
        <div class="summary-card" style="--delay: 140ms">
          <CheckCircle :size="22" />
          <strong>{{ summary.presentCount }}</strong>
          <span>Present At Least Once</span>
          <small>Met stay threshold</small>
        </div>
        <div class="summary-card" style="--delay: 210ms">
          <XCircle :size="22" />
          <strong>{{ summary.absentCount }}</strong>
          <span>Fully Absent</span>
          <small>No qualifying sessions</small>
        </div>
        <div class="summary-card" style="--delay: 280ms">
          <Sparkles :size="22" />
          <strong>{{ summary.unmatchedRecordsCount }}</strong>
          <span>Unmatched Records</span>
          <small>Outside nomination list</small>
        </div>
      </section>

      <section class="dashboard-card preview-card">
        <div class="section-heading">
          <div>
            <h2>Preview of extracted report</h2>
            <p>{{ exportReady ? 'Showing the first five rows from the generated report.' : 'Extract attendance to preview the report structure.' }}</p>
          </div>
        </div>
        <div v-if="previewRows.length" class="preview-table-wrap">
          <table class="preview-table">
            <thead>
              <tr>
                <th>{{ attendanceType === 'Webex' ? 'Mail_Id' : 'Emp_Id' }}</th>
                <th>Emp_Name</th>
                <th>{{ dateColumns[0] || 'Date' }}</th>
                <th>Duration</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="row in previewRows" :key="`${row.id}-${row.name}`">
                <td>{{ row.id }}</td>
                <td>{{ row.name }}</td>
                <td><span :class="['status-badge', row.status.toLowerCase()]">{{ row.status }}</span></td>
                <td>{{ row.duration }}</td>
              </tr>
            </tbody>
          </table>
        </div>
        <div v-else class="empty-state">
          <FileSpreadsheet :size="30" />
          <strong>No preview yet</strong>
          <span>Your extracted rows will appear here after processing.</span>
        </div>
      </section>
    </main>
  </div>
</template>

<script>
import ExcelJS from 'exceljs';
import Papa from 'papaparse';
import moment from 'moment';
import confetti from 'canvas-confetti';
import {
  CheckCircle,
  Download,
  FileCheck,
  FileSpreadsheet,
  Mail,
  RotateCcw,
  Sparkles,
  Timer,
  Upload,
  Users,
  XCircle,
} from 'lucide-vue-next';

export default {
  name: 'AttendanceReportGenerator',
  components: {
    CheckCircle,
    Download,
    FileCheck,
    FileSpreadsheet,
    Mail,
    RotateCcw,
    Sparkles,
    Timer,
    Upload,
    Users,
    XCircle,
  },
  data() {
    return {
      attendanceType: '',
      attendanceFiles: [],
      nominationOption: '',
      nominationFile: null,
      minDurationStay: 0,
      sessions: [],
      finalAttendance: [],
      unmatchedRecords: [],
      summary: {
        totalNominatedEmployees: 0,
        totalUploadedAttendanceFiles: 0,
        presentCount: 0,
        absentCount: 0,
        unmatchedRecordsCount: 0,
      },
      loading: false,
      exportReady: false,
      errorMessage: '',
      successMessage: '',
      nominatedEmployees: [],
    };
  },
  computed: {
    canExtract() {
      return (
        this.attendanceType &&
        this.attendanceFiles.length > 0 &&
        this.nominationOption &&
        (this.nominationOption === 'No' || this.nominationFile)
      );
    },
    nominationHint() {
      return this.attendanceType === 'Webex'
        ? 'Expected columns: Emp Name, Emp Mail ID'
        : 'Expected columns: Emp ID, Emp Name';
    },
    activeStep() {
      if (this.exportReady) return 5;
      if (this.loading) return 4;
      if (this.nominationOption && (this.nominationOption === 'No' || this.nominationFile)) return 4;
      if (this.attendanceFiles.length) return 3;
      if (this.attendanceType) return 2;
      return 1;
    },
    selectedTypeLabel() {
      if (!this.attendanceType) return '';
      return this.attendanceType === 'Teams' ? 'Microsoft Teams selected' : 'Cisco Webex selected';
    },
    selectedFilesCount() {
      return this.attendanceFiles.length;
    },
    dateColumns() {
      return this.sessions.map((session) => session.date);
    },
    previewRows() {
      if (!this.exportReady || !this.finalAttendance.length) return [];
      const firstDate = this.dateColumns[0];
      return this.finalAttendance.slice(0, 5).map((employee) => ({
        id: this.attendanceType === 'Webex' ? this.getWebexExportEmail(employee.EMAIL) : employee.EMPID || '',
        name: employee.NAME || '',
        status: employee.Attendance[firstDate] || 'Absent',
        duration: employee.Duration[firstDate] || '0 mins',
      }));
    },
    stepItems() {
      return [
        { id: 1, label: 'Select Source', icon: Sparkles },
        { id: 2, label: 'Upload Files', icon: Upload },
        { id: 3, label: 'Nominations', icon: Users },
        { id: 4, label: 'Extract', icon: FileCheck },
        { id: 5, label: 'Export', icon: Download },
      ];
    },
  },
  methods: {
    handleAttendanceTypeSelection(type) {
      if (this.attendanceType && this.attendanceType !== type) {
        this.attendanceFiles = [];
        this.sessions = [];
        this.finalAttendance = [];
        this.unmatchedRecords = [];
        this.exportReady = false;
      }
      this.attendanceType = type;
    },

    handleAttendanceDrop(event, type) {
      const files = Array.from(event.dataTransfer.files || []);
      this.setAttendanceFiles(files, type);
    },

    handleAttendanceUpload(event, type) {
      const files = Array.from(event.target.files || []);
      this.setAttendanceFiles(files, type, event.target);
    },

    setAttendanceFiles(files, type, inputElement = null) {
      this.clearMessages();
      if (!files.length) return;

      const invalidFile = files.find((file) => !this.isCsvFile(file));
      if (invalidFile) {
        this.showError(`"${invalidFile.name}" is not a valid CSV file.`);
        if (inputElement) inputElement.value = '';
        return;
      }

      this.handleAttendanceTypeSelection(type);
      this.attendanceFiles = files;
      this.sessions = [];
      this.finalAttendance = [];
      this.unmatchedRecords = [];
      this.exportReady = false;

      if (type === 'Teams' && this.$refs.webexInput) this.$refs.webexInput.value = '';
      if (type === 'Webex' && this.$refs.teamsInput) this.$refs.teamsInput.value = '';
    },

    removeAttendanceFile(index) {
      this.attendanceFiles = this.attendanceFiles.filter((_, fileIndex) => fileIndex !== index);
      this.sessions = [];
      this.finalAttendance = [];
      this.unmatchedRecords = [];
      this.exportReady = false;
      this.clearMessages();
      if (!this.attendanceFiles.length) {
        if (this.$refs.teamsInput) this.$refs.teamsInput.value = '';
        if (this.$refs.webexInput) this.$refs.webexInput.value = '';
      }
    },

    removeNominationFile() {
      this.nominationFile = null;
      this.exportReady = false;
      this.clearMessages();
      if (this.$refs.nominationInput) this.$refs.nominationInput.value = '';
    },

    formatFileSize(bytes) {
      const size = Number(bytes || 0);
      if (size < 1024) return `${size} B`;
      if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
      return `${(size / (1024 * 1024)).toFixed(1)} MB`;
    },

    handleNominationOptionChange(option) {
      this.clearMessages();
      this.nominationOption = option;
      this.nominationFile = null;
      this.nominatedEmployees = [];
      this.finalAttendance = [];
      this.unmatchedRecords = [];
      this.exportReady = false;
      if (this.$refs.nominationInput) this.$refs.nominationInput.value = '';
    },

    handleNominationUpload(event) {
      this.clearMessages();
      const file = event.target.files && event.target.files[0];
      if (!file) return;

      if (!/\.(xlsx|xls)$/i.test(file.name)) {
        this.showError('Invalid nomination file. Please upload an Excel file with .xlsx or .xls extension.');
        event.target.value = '';
        return;
      }

      this.nominationFile = file;
      this.exportReady = false;
    },

    async extractTeamsAttendance(file) {
      const rows = await this.parseTeamsCsvFile(file);
      const date = this.extractTeamsDateFromFilename(file.name);
      const participantStartIndex = rows.findIndex((row) =>
        this.getTeamsRowCells(row).some((cell) => /^2\.\s*participants$/i.test(cell))
      );

      if (participantStartIndex === -1) {
        throw new Error(`Invalid Teams attendance file. Could not find the "2. Participants" section in "${file.name}".`);
      }

      const participants = [];
      let participantHeaders = [];

      for (let index = participantStartIndex + 1; index < rows.length; index += 1) {
        const cells = this.getTeamsRowCells(rows[index]);
        if (!cells.length) continue;

        const joined = cells.join(' ').trim();
        if (/^\d+\.\s+/.test(joined) && !/^2\.\s*participants$/i.test(joined)) break;

        const lowerCells = cells.map((cell) => cell.toLowerCase());
        if (this.isTeamsParticipantHeader(lowerCells)) {
          participantHeaders = lowerCells;
          continue;
        }

        const participant = this.extractTeamsParticipantFromCells(cells, participantHeaders, date);
        if (participant && (participant.NAME || participant.EMAIL || participant.EMPID)) {
          participants.push(participant);
        }
      }

      const mergedParticipants = this.mergeDuplicateParticipants(participants, 'Teams');
      if (!mergedParticipants.length) {
        throw new Error(`No participants found in "${file.name}".`);
      }

      return {
        date,
        source: 'Teams',
        participants: mergedParticipants,
      };
    },

    async extractWebexAttendance(file) {
      const rows = await this.parseCsvFile(file, 'Webex');
      const normalizedRows = rows.map((row) => this.normalizeRow(row)).filter((row) => Object.keys(row).length);

      if (!this.hasWebexColumns(normalizedRows)) {
        throw new Error(`Invalid Webex attendance file. Expected Meeting Start Time, Display Name, Attendee Email, and Attendance Duration columns in "${file.name}".`);
      }

      let sessionDate = '';
      const participants = [];

      normalizedRows.forEach((row) => {
        const meetingStartTime = this.pickValue(row, ['meeting start time', 'start time', 'meeting date']);
        const displayName = this.pickValue(row, ['display name', 'name', 'attendee name']);
        const email = this.pickValue(row, ['attendee email', 'email', 'email address']);
        const duration = this.pickValue(row, ['attendance duration', 'duration', 'attended duration']);

        if (!sessionDate && meetingStartTime) sessionDate = this.extractWebexDate(meetingStartTime, file.name);
        if (!displayName && !email) return;

        participants.push({
          EMPID: '',
          NAME: displayName || email,
          EMAIL: email || '',
          RAW_EMAIL: email || '',
          DURATION: duration || '0 mins',
          DURATION_MINUTES: this.convertDurationToMinutes(duration),
          DATE: '',
        });
      });

      const mergedParticipants = this.mergeDuplicateParticipants(participants, 'Webex').map((participant) => ({
        ...participant,
        DATE: sessionDate,
      }));

      if (!mergedParticipants.length) {
        throw new Error(`No participants found in "${file.name}".`);
      }

      return {
        date: sessionDate || this.extractWebexDate('', file.name),
        source: 'Webex',
        participants: mergedParticipants,
      };
    },

    async extractNomination() {
      if (this.nominationOption !== 'Yes') return [];

      try {
        const workbook = new ExcelJS.Workbook();
        const buffer = await this.nominationFile.arrayBuffer();
        await workbook.xlsx.load(buffer);

        const worksheet = workbook.worksheets[0];
        if (!worksheet) throw new Error('Invalid nomination file. No worksheet was found.');

        const nominations = [];
        worksheet.eachRow((row, rowNumber) => {
          const first = this.cellText(row.getCell(1));
          const second = this.cellText(row.getCell(2));
          if (!first && !second) return;
          if (rowNumber === 1 && this.looksLikeHeader(first, second)) return;

          if (this.attendanceType === 'Teams') {
            if (!first || !second) return;
            nominations.push({
              EMPID: first,
              NAME: second,
              EMAIL: '',
              Attendance: {},
              Duration: {},
            });
          } else {
            if (!first || !second || !this.looksLikeEmail(second)) return;
            nominations.push({
              EMPID: '',
              NAME: first,
              EMAIL: second,
              Attendance: {},
              Duration: {},
            });
          }
        });

        if (!nominations.length) {
          throw new Error(
            this.attendanceType === 'Teams'
              ? 'Invalid nomination file. Teams nomination must have Emp ID in column 1 and Emp Name in column 2.'
              : 'Invalid nomination file. Webex nomination must have Emp Name in column 1 and Emp Mail ID in column 2.'
          );
        }

        return this.mergeNominations(nominations);
      } catch (error) {
        if (error.message && error.message.startsWith('Invalid nomination file')) throw error;
        throw new Error('Invalid nomination file. Please check the workbook format.');
      }
    },

    async prepareFinalAttendance() {
      if (!this.canExtract) return;

      this.loading = true;
      this.clearMessages();
      this.sessions = [];
      this.finalAttendance = [];
      this.unmatchedRecords = [];
      this.exportReady = false;

      try {
        const sessionPromises = this.attendanceFiles.map((file) =>
          this.attendanceType === 'Teams' ? this.extractTeamsAttendance(file) : this.extractWebexAttendance(file)
        );
        const extractedSessions = await Promise.all(sessionPromises);
        this.sessions = this.assignUniqueSortedDates(extractedSessions).map((session) => ({
          ...session,
          participants: this.mergeDuplicateParticipants(session.participants, session.source),
        }));
        this.nominatedEmployees = await this.extractNomination();

        const sessionDates = this.sessions.map((session) => session.date);
        const employeeMap = new Map();

        if (this.nominationOption === 'Yes') {
          this.nominatedEmployees.forEach((employee) => {
            const key = this.buildParticipantKey(employee, this.attendanceType);
            employeeMap.set(key, {
              EMPID: employee.EMPID || '',
              NAME: employee.NAME || '',
              EMAIL: employee.EMAIL || '',
              Attendance: this.emptyAttendance(sessionDates),
              Duration: this.emptyDuration(sessionDates),
            });
          });
        } else {
          this.sessions.forEach((session) => {
            session.participants.forEach((participant) => {
              const key = this.buildParticipantKey(participant, session.source);
              if (!key || employeeMap.has(key)) return;
              employeeMap.set(key, {
                EMPID: participant.EMPID || '',
                NAME: participant.NAME || '',
                EMAIL: participant.EMAIL || '',
                Attendance: this.emptyAttendance(sessionDates),
                Duration: this.emptyDuration(sessionDates),
              });
            });
          });
        }

        this.sessions.forEach((session) => {
          const participantMap = new Map(session.participants.map((participant) => [this.buildParticipantKey(participant, session.source), participant]));
          const webexEmailMap = new Map();
          if (session.source === 'Webex') {
            session.participants.forEach((participant) => {
              const emailKey = this.normalizeEmail(participant.EMAIL);
              if (this.isRealEmail(emailKey)) webexEmailMap.set(emailKey, participant);
            });
          }

          employeeMap.forEach((employee, key) => {
            const participant =
              participantMap.get(key) ||
              (session.source === 'Webex' ? webexEmailMap.get(this.normalizeEmail(employee.EMAIL)) : null);
            const durationMinutes = participant ? Number(participant.DURATION_MINUTES || 0) : 0;
            employee.Attendance[session.date] = durationMinutes >= Number(this.minDurationStay) ? 'Present' : 'Absent';
            employee.Duration[session.date] = this.formatDurationFromMinutes(durationMinutes);
            if (participant && this.attendanceType === 'Webex' && !this.isRealEmail(employee.EMAIL)) {
              employee.EMAIL = this.getPreferredWebexEmail(employee.EMAIL, participant.EMAIL);
            } else if (participant && !employee.EMAIL) {
              employee.EMAIL = participant.EMAIL || '';
            }
            if (participant && !employee.EMPID) employee.EMPID = participant.EMPID || '';
            if (participant && !employee.NAME) employee.NAME = participant.NAME || '';
          });
        });

        this.finalAttendance = Array.from(employeeMap.values());
        this.unmatchedRecords = this.findUnmatchedRecords();
        this.updateSummary();
        this.exportReady = true;
        this.successMessage = 'Attendance extracted successfully. You can export the Excel report now.';
      } catch (error) {
        this.showError(error.message || 'Unable to extract attendance. Please check the uploaded files.');
      } finally {
        this.loading = false;
      }
    },

    findUnmatchedRecords() {
      if (this.nominationOption !== 'Yes') return [];

      const nominatedKeys = new Set(this.nominatedEmployees.map((employee) => this.buildParticipantKey(employee, this.attendanceType)));
      const nominatedEmailKeys = new Set(this.nominatedEmployees.map((employee) => this.normalizeEmail(employee.EMAIL)).filter(Boolean));
      const unmatchedMap = new Map();

      this.sessions.forEach((session) => {
        session.participants.forEach((participant) => {
          const key = this.buildParticipantKey(participant, session.source);
          const emailKey = this.normalizeEmail(participant.EMAIL);
          const isMatchedByEmail = session.source === 'Webex' && this.isRealEmail(emailKey) && nominatedEmailKeys.has(emailKey);
          if (!key || nominatedKeys.has(key) || isMatchedByEmail) return;

          const recordKey = `${session.date}|${key}`;
          const existing = unmatchedMap.get(recordKey);
          if (existing) {
            existing.DURATION_MINUTES += Number(participant.DURATION_MINUTES || 0);
            existing.Duration = this.formatDurationFromMinutes(existing.DURATION_MINUTES);
            if (session.source === 'Webex') existing.EMAIL = this.getPreferredWebexEmail(existing.EMAIL, participant.EMAIL);
            return;
          }

          unmatchedMap.set(recordKey, {
            EMPID: participant.EMPID || '',
            NAME: participant.NAME || '',
            EMAIL: session.source === 'Webex' ? this.getPreferredWebexEmail('', participant.EMAIL) : participant.EMAIL || '',
            Date: session.date,
            Duration: this.formatDurationFromMinutes(participant.DURATION_MINUTES || 0),
            DURATION_MINUTES: Number(participant.DURATION_MINUTES || 0),
            Source: session.source,
          });
        });
      });

      return Array.from(unmatchedMap.values());
    },

    filterParticipantsByMinimumStay(participants) {
      return participants.filter((participant) => Number(participant.DURATION_MINUTES || 0) >= Number(this.minDurationStay));
    },

    normalizeText(value) {
      return String(value || '').trim().replace(/\./g, '').replace(/\s+/g, ' ').toLowerCase();
    },

    normalizeEmail(value) {
      return String(value || '').trim().toLowerCase();
    },

    normalizeEmpId(value) {
      return String(value || '').trim().toLowerCase();
    },

    isGuestWebexEmail(value) {
      return this.normalizeEmail(value).includes('guest.webex.localhost');
    },

    isRealEmail(value) {
      const email = this.normalizeEmail(value);
      return email.includes('@') && !this.isGuestWebexEmail(email);
    },

    getPreferredWebexEmail(currentEmail, nextEmail) {
      if (this.isRealEmail(currentEmail)) return this.normalizeEmail(currentEmail);
      if (this.isRealEmail(nextEmail)) return this.normalizeEmail(nextEmail);
      if (this.normalizeEmail(currentEmail) === 'guest' || this.isGuestWebexEmail(currentEmail)) return 'guest';
      if (this.normalizeEmail(nextEmail) === 'guest' || this.isGuestWebexEmail(nextEmail)) return 'guest';
      return this.normalizeEmail(currentEmail || nextEmail) || 'guest';
    },

    getWebexExportEmail(email) {
      return this.isRealEmail(email) ? this.normalizeEmail(email) : 'guest';
    },

    buildParticipantKey(participant, source = this.attendanceType) {
      if (source === 'Teams') {
        return this.normalizeEmpId(participant.EMPID) || this.normalizeText(participant.NAME);
      }
      return this.normalizeText(participant.NAME) || this.normalizeEmail(participant.EMAIL);
    },

    mergeDuplicateParticipants(participants, source = this.attendanceType) {
      const participantMap = new Map();

      participants.forEach((participant) => {
        const key = this.buildParticipantKey(participant, source);
        if (!key) return;

        const durationMinutes = Number(participant.DURATION_MINUTES || 0);
        const existing = participantMap.get(key);
        if (existing) {
          existing.DURATION_MINUTES += durationMinutes;
          existing.DURATION = this.formatDurationFromMinutes(existing.DURATION_MINUTES);
          if (source === 'Webex') {
            existing.EMAIL = this.getPreferredWebexEmail(existing.EMAIL, participant.EMAIL);
          } else if (!existing.EMAIL) {
            existing.EMAIL = participant.EMAIL || '';
          }
          if (!existing.EMPID) existing.EMPID = participant.EMPID || '';
          if (!existing.NAME) existing.NAME = participant.NAME || '';
          return;
        }

        const email = source === 'Webex' ? this.getPreferredWebexEmail('', participant.EMAIL) : participant.EMAIL || '';
        participantMap.set(key, {
          EMPID: participant.EMPID || '',
          NAME: participant.NAME || '',
          EMAIL: email,
          RAW_EMAIL: participant.RAW_EMAIL || participant.EMAIL || '',
          DURATION: this.formatDurationFromMinutes(durationMinutes),
          DURATION_MINUTES: durationMinutes,
          DATE: participant.DATE || '',
        });
      });

      return Array.from(participantMap.values());
    },

    mergeNominations(nominations) {
      const map = new Map();
      nominations.forEach((employee) => {
        const key = this.buildParticipantKey(employee, this.attendanceType);
        if (key && !map.has(key)) map.set(key, employee);
      });
      return Array.from(map.values());
    },

    convertDurationToMinutes(value) {
      if (value === null || value === undefined) return 0;
      if (typeof value === 'number') return Math.round(value);

      const text = String(value).trim().toLowerCase();
      if (!text) return 0;
      if (/^\d+(\.\d+)?$/.test(text)) return Math.round(Number(text));

      const timeMatch = text.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
      if (timeMatch) {
        const first = Number(timeMatch[1]);
        const second = Number(timeMatch[2]);
        const third = Number(timeMatch[3] || 0);
        return timeMatch[3] ? first * 60 + second + Math.ceil(third / 60) : first * 60 + second;
      }

      let minutes = 0;
      const hours = text.match(/(\d+(?:\.\d+)?)\s*(h|hr|hrs|hour|hours)/);
      const mins = text.match(/(\d+(?:\.\d+)?)\s*(m|min|mins|minute|minutes)/);
      const secs = text.match(/(\d+(?:\.\d+)?)\s*(s|sec|secs|second|seconds)/);

      if (hours) minutes += Number(hours[1]) * 60;
      if (mins) minutes += Number(mins[1]);
      if (secs) minutes += Number(secs[1]) / 60;

      return Math.ceil(minutes);
    },

    formatDurationFromMinutes(totalMinutes) {
      const mins = Math.round(Number(totalMinutes) || 0);
      const hours = Math.floor(mins / 60);
      const remainingMinutes = mins % 60;

      if (hours === 0) return `${remainingMinutes} mins`;
      if (remainingMinutes === 0) return `${hours}hr`;
      return `${hours}hr ${remainingMinutes} mins`;
    },

    formatMinutes(value) {
      return this.formatDurationFromMinutes(value);
    },

    extractTeamsDateFromFilename(filename) {
      const cleanName = filename.replace(/\.[^.]+$/, '');
      const dateMatch = cleanName.match(/(\d{1,2})[-_/](\d{1,2})[-_/](\d{2,4})/);
      if (dateMatch) {
        return moment(dateMatch[0], ['M-D-YY', 'M-D-YYYY', 'M/D/YY', 'M/D/YYYY', 'M_D_YY', 'M_D_YYYY']).format('DD-MMM-YY');
      }
      return moment().format('DD-MMM-YY');
    },

    extractDateFromFilename(filename) {
      const cleanName = String(filename || '').replace(/\.[^.]+$/, '');
      const dateMatch = cleanName.match(/(\d{1,2})[-_/](\d{1,2})[-_/](\d{2,4})/);
      if (dateMatch) {
        const parsed = moment(dateMatch[0], ['M-D-YY', 'M-D-YYYY', 'M/D/YY', 'M/D/YYYY', 'M_D_YY', 'M_D_YYYY'], true);
        if (parsed.isValid()) return parsed.format('DD-MMM-YY');
      }
      return moment().format('DD-MMM-YY');
    },

    extractWebexDate(value, filename = '') {
      const cleaned = String(value || '')
        .replace(/^="/, '')
        .replace(/^=/, '')
        .replace(/"$/g, '')
        .replace(/"/g, '')
        .trim();
      const parsed = moment(cleaned, [
        moment.ISO_8601,
        'M/D/YYYY, h:mm:ss A',
        'M/D/YYYY h:mm:ss A',
        'MM/DD/YYYY, h:mm:ss A',
        'MM/DD/YYYY h:mm:ss A',
        'M/D/YY, h:mm:ss A',
        'M/D/YY h:mm:ss A',
        'YYYY-MM-DD HH:mm:ss',
      ], true);

      if (parsed.isValid()) return parsed.format('DD-MMM-YY');
      return this.extractDateFromFilename(filename);
    },

    async exportExcel() {
      if (!this.exportReady) return;

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Attendance Report');
      const dateColumns = this.sessions.map((session) => session.date);
      const idHeader = this.attendanceType === 'Webex' ? 'Mail_Id' : 'Emp_Id';
      const header = [idHeader, 'Emp_Name'];

      dateColumns.forEach((date) => {
        header.push(date, 'Duration');
      });

      worksheet.addRow(header);
      this.finalAttendance.forEach((employee) => {
        const row = [this.attendanceType === 'Webex' ? this.getWebexExportEmail(employee.EMAIL) : employee.EMPID || '', employee.NAME || ''];
        dateColumns.forEach((date) => {
          row.push(employee.Attendance[date] || 'Absent', employee.Duration[date] || '0 mins');
        });
        worksheet.addRow(row);
      });

      worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: header.length },
      };
      worksheet.views = [{ state: 'frozen', ySplit: 1 }];
      this.styleMainWorksheet(worksheet, header.length);

      if (this.nominationOption === 'Yes') {
        worksheet.addRow([]);
        const titleRow = worksheet.addRow(['These are the unmatched records']);
        titleRow.font = { bold: true, color: { argb: 'FF111827' } };
        titleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF59E0B' } };
        worksheet.mergeCells(titleRow.number, 1, titleRow.number, 6);
        titleRow.eachCell((cell) => this.applyBorder(cell));

        const unmatchedIdHeader = this.attendanceType === 'Webex' ? 'Mail_Id' : 'Emp_Id';
        if (this.unmatchedRecords.length) {
          const unmatchedHeader = worksheet.addRow([unmatchedIdHeader, 'Emp_Name', 'Email', 'Date', 'Duration', 'Source']);
          unmatchedHeader.font = { bold: true, color: { argb: 'FF111827' } };
          unmatchedHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBBF24' } };
          unmatchedHeader.eachCell((cell) => {
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            this.applyBorder(cell);
          });

          this.unmatchedRecords.forEach((record) => {
            const row = worksheet.addRow([
              this.attendanceType === 'Webex' ? this.getWebexExportEmail(record.EMAIL) : record.EMPID,
              record.NAME,
              this.attendanceType === 'Webex' ? this.getWebexExportEmail(record.EMAIL) : record.EMAIL,
              record.Date,
              record.Duration,
              record.Source,
            ]);
            row.eachCell((cell, colNumber) => {
              cell.alignment = { horizontal: colNumber <= 3 ? 'left' : 'center', vertical: 'middle' };
              this.applyBorder(cell);
            });
          });
        } else {
          const noRecordRow = worksheet.addRow(['No unmatched records found']);
          worksheet.mergeCells(noRecordRow.number, 1, noRecordRow.number, 6);
          noRecordRow.eachCell((cell) => this.applyBorder(cell));
        }
      }

      this.autoFitColumns(worksheet);

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `Attendance_Report_${moment().format('YYYYMMDD_HHmm')}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
      this.successMessage = 'Excel report exported successfully.';
      confetti({
        particleCount: 90,
        spread: 62,
        origin: { y: 0.72 },
        colors: ['#2563eb', '#06b6d4', '#10b981', '#f59e0b'],
      });
    },

    styleMainWorksheet(worksheet, headerLength) {
      const mainHeader = worksheet.getRow(1);
      mainHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      mainHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
      mainHeader.alignment = { horizontal: 'center', vertical: 'middle' };

      worksheet.eachRow((row, rowNumber) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          if (colNumber > headerLength) return;
          this.applyBorder(cell);
          if (rowNumber === 1) return;

          if (colNumber <= 2) {
            cell.alignment = { horizontal: 'left', vertical: 'middle' };
            return;
          }

          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          const isDurationColumn = colNumber >= 4 && colNumber % 2 === 0;
          if (isDurationColumn) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
          } else if (cell.value === 'Present') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
          } else if (cell.value === 'Absent') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
          }
        });
      });
    },

    resetForm() {
      this.attendanceType = '';
      this.attendanceFiles = [];
      this.nominationOption = '';
      this.nominationFile = null;
      this.minDurationStay = 0;
      this.sessions = [];
      this.finalAttendance = [];
      this.unmatchedRecords = [];
      this.nominatedEmployees = [];
      this.exportReady = false;
      this.loading = false;
      this.clearMessages();
      this.summary = {
        totalNominatedEmployees: 0,
        totalUploadedAttendanceFiles: 0,
        presentCount: 0,
        absentCount: 0,
        unmatchedRecordsCount: 0,
      };

      if (this.$refs.teamsInput) this.$refs.teamsInput.value = '';
      if (this.$refs.webexInput) this.$refs.webexInput.value = '';
      if (this.$refs.nominationInput) this.$refs.nominationInput.value = '';
    },

    parseTeamsCsvFile(file) {
      return new Promise((resolve, reject) => {
        Papa.parse(file, {
          header: true,
          skipEmptyLines: false,
          dynamicTyping: false,
          transformHeader: (header) => String(header || '').trim(),
          complete: (result) => {
            resolve(result.data || []);
          },
          error: () => reject(new Error(`Invalid Teams attendance file. Unable to read "${file.name}".`)),
        });
      });
    },

    parseCsvFile(file) {
      return new Promise((resolve, reject) => {
        Papa.parse(file, {
          header: true,
          skipEmptyLines: true,
          dynamicTyping: false,
          transformHeader: (header) => String(header || '').trim(),
          complete: (result) => resolve(result.data || []),
          error: () => reject(new Error(`Invalid Webex attendance file. Unable to read "${file.name}".`)),
        });
      });
    },

    getTeamsRowCells(row) {
      const rawValues = [];

      Object.entries(row || {}).forEach(([key, value]) => {
        if (key && key !== '__parsed_extra' && this.shouldKeepTeamsHeaderKey(key)) rawValues.push(key);
        if (key === '__parsed_extra' && Array.isArray(value)) {
          rawValues.push(...value);
          return;
        }
        rawValues.push(value);
      });

      return rawValues
        .flatMap((value) => String(value ?? '').split('\t'))
        .map((value) => value.replace(/^"|"$/g, '').trim())
        .filter(Boolean);
    },

    shouldKeepTeamsHeaderKey(key) {
      const value = String(key || '').trim().toLowerCase();
      return /^2\.\s*participants$/.test(value) || /^name$/.test(value) || value.includes('email') || value.includes('duration');
    },

    isTeamsParticipantHeader(lowerCells) {
      const joined = lowerCells.join(' ');
      return (
        lowerCells.includes('name') ||
        joined.includes('full name') ||
        joined.includes('participant')
      ) && (
        joined.includes('email') ||
        joined.includes('duration') ||
        joined.includes('in-meeting')
      );
    },

    extractTeamsParticipantFromCells(cells, lowerHeaders, date) {
      const rowByHeader = this.mapTeamsCellsByHeader(cells, lowerHeaders);
      const email = this.findTeamsEmail(cells, rowByHeader);
      const duration = this.findTeamsDuration(cells, rowByHeader);
      const name =
        rowByHeader.name ||
        rowByHeader['full name'] ||
        rowByHeader['display name'] ||
        rowByHeader['participant name'] ||
        rowByHeader['user name'] ||
        this.findTeamsName(cells, email, duration);

      if (!name && !email) return null;

      const durationMinutes = this.convertDurationToMinutes(duration);
      const empId = this.extractEmpIdFromEmail(email) || this.extractEmpIdFromText(name);

      return {
        EMPID: empId || name || email,
        NAME: name || email || empId,
        EMAIL: email || '',
        DURATION: this.formatMinutes(durationMinutes),
        DURATION_MINUTES: durationMinutes,
        DATE: date,
      };
    },

    mapTeamsCellsByHeader(cells, lowerHeaders) {
      if (!lowerHeaders.length) return {};

      return lowerHeaders.reduce((mapped, header, index) => {
        const normalizedHeader = this.normalizeTeamsHeader(header);
        if (normalizedHeader && cells[index] !== undefined) {
          mapped[normalizedHeader] = cells[index];
        }
        return mapped;
      }, {});
    },

    normalizeTeamsHeader(header) {
      const value = String(header || '').trim().toLowerCase();
      if (['name', 'full name', 'display name', 'participant name', 'user name'].includes(value)) return value;
      if (value.includes('email') || value.includes('upn')) return 'email';
      if (value.includes('duration') || value.includes('in-meeting')) return 'duration';
      return value;
    },

    findTeamsEmail(cells, rowByHeader) {
      if (rowByHeader.email) return rowByHeader.email;
      return cells.find((cell) => this.looksLikeEmail(cell)) || '';
    },

    findTeamsDuration(cells, rowByHeader) {
      if (rowByHeader.duration) return rowByHeader.duration;
      return (
        cells.find((cell) =>
          /(\d+\s*(h|hr|hrs|hour|hours|min|mins|minute|minutes|sec|secs|second|seconds)\b)|(^\d{1,2}:\d{1,2}(:\d{1,2})?$)/i.test(cell)
        ) || ''
      );
    },

    findTeamsName(cells, email, duration) {
      const ignored = new Set([email, duration].filter(Boolean));
      return (
        cells.find((cell) => {
          if (ignored.has(cell)) return false;
          if (this.looksLikeEmail(cell)) return false;
          if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/.test(cell)) return false;
          if (/^\d{1,2}:\d{2}/.test(cell)) return false;
          if (/^(organizer|presenter|attendee)$/i.test(cell)) return false;
          return /[a-z]/i.test(cell);
        }) || ''
      );
    },

    hasWebexColumns(rows) {
      if (!rows.length) return false;
      const keys = new Set(Object.keys(rows[0] || {}));
      const hasStart = ['meeting start time', 'start time', 'meeting date'].some((key) => keys.has(key));
      const hasName = ['display name', 'name', 'attendee name'].some((key) => keys.has(key));
      const hasEmail = ['attendee email', 'email', 'email address'].some((key) => keys.has(key));
      const hasDuration = ['attendance duration', 'duration', 'attended duration'].some((key) => keys.has(key));
      return hasStart && hasName && hasEmail && hasDuration;
    },

    assignUniqueSortedDates(sessions) {
      const sorted = [...sessions].sort((a, b) => {
        const dateA = moment(a.date, 'DD-MMM-YY');
        const dateB = moment(b.date, 'DD-MMM-YY');
        return dateA.valueOf() - dateB.valueOf();
      });
      const counts = {};

      return sorted.map((session) => {
        const baseDate = session.date;
        counts[baseDate] = (counts[baseDate] || 0) + 1;
        const uniqueDate = counts[baseDate] === 1 ? baseDate : `${baseDate} (${counts[baseDate]})`;
        return {
          ...session,
          date: uniqueDate,
          participants: session.participants.map((participant) => ({
            ...participant,
            DATE: uniqueDate,
          })),
        };
      });
    },

    getEmployeeKey(employee) {
      return this.buildParticipantKey(employee, this.attendanceType);
    },

    emptyAttendance(dates) {
      return dates.reduce((attendance, date) => {
        attendance[date] = 'Absent';
        return attendance;
      }, {});
    },

    emptyDuration(dates) {
      return dates.reduce((duration, date) => {
        duration[date] = '0 mins';
        return duration;
      }, {});
    },

    updateSummary() {
      const presentCount = this.finalAttendance.filter((employee) =>
        Object.values(employee.Attendance).some((status) => status === 'Present')
      ).length;

      this.summary = {
        totalNominatedEmployees: this.finalAttendance.length,
        totalUploadedAttendanceFiles: this.attendanceFiles.length,
        presentCount,
        absentCount: this.finalAttendance.length - presentCount,
        unmatchedRecordsCount: this.unmatchedRecords.length,
      };
    },

    normalizeRow(row) {
      return Object.entries(row || {}).reduce((normalized, [key, value]) => {
        const cleanKey = String(key || '').trim().toLowerCase();
        if (cleanKey) normalized[cleanKey] = typeof value === 'string' ? value.trim() : value;
        return normalized;
      }, {});
    },

    pickValue(row, possibleKeys) {
      const foundKey = possibleKeys.find((key) => Object.prototype.hasOwnProperty.call(row, key));
      return foundKey ? String(row[foundKey] || '').trim() : '';
    },

    cellText(cell) {
      if (!cell || cell.value === null || cell.value === undefined) return '';
      if (typeof cell.value === 'object') {
        if (cell.value.text) return String(cell.value.text).trim();
        if (cell.value.result) return String(cell.value.result).trim();
        if (cell.value.richText) return cell.value.richText.map((item) => item.text).join('').trim();
      }
      return String(cell.value).trim();
    },

    extractEmpIdFromEmail(email) {
      const value = String(email || '').trim().toLowerCase();
      if (!value.includes('@hexaware.com')) return '';
      const localPart = value.split('@')[0];
      const digitMatch = localPart.match(/\d{4,}/);
      return digitMatch ? digitMatch[0] : localPart;
    },

    extractEmpIdFromText(text) {
      const match = String(text || '').match(/\b\d{4,}\b/);
      return match ? match[0] : '';
    },

    looksLikeHeader(first, second) {
      const text = `${first} ${second}`.toLowerCase();
      return /emp|employee|name|mail|email|id/.test(text);
    },

    looksLikeEmail(value) {
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
    },

    isCsvFile(file) {
      return /\.csv$/i.test(file.name) || file.type === 'text/csv';
    },

    applyBorder(cell) {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        right: { style: 'thin', color: { argb: 'FFD1D5DB' } },
      };
    },

    autoFitColumns(worksheet) {
      worksheet.columns.forEach((column) => {
        let maxLength = 12;
        column.eachCell({ includeEmpty: true }, (cell) => {
          const value = cell.value ? String(cell.value) : '';
          maxLength = Math.max(maxLength, value.length + 2);
        });
        column.width = Math.min(maxLength, 38);
      });
    },

    showError(message) {
      this.errorMessage = message;
      this.successMessage = '';
    },

    clearMessages() {
      this.errorMessage = '';
      this.successMessage = '';
    },
  },
};
</script>

<style scoped>
.attendance-page {
  --primary: #273a8a;
  --primary-strong: #172554;
  --accent: #06b6d4;
  --success: #10b981;
  --warning: #f59e0b;
  --danger: #e11d48;
  --ink: #0f172a;
  --muted: #64748b;
  --line: rgba(148, 163, 184, 0.32);
  --glass: rgba(255, 255, 255, 0.74);
  --shadow: 0 22px 55px rgba(15, 23, 42, 0.12);
  min-height: 100vh;
  position: relative;
  overflow: hidden;
  padding: 42px 24px 54px;
  background: linear-gradient(120deg, #eef6ff, #f8fbff, #eef2ff, #ecfeff);
  background-size: 240% 240%;
  animation: gradientShift 16s ease infinite;
  color: var(--ink);
  font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}

.background-grid {
  position: fixed;
  inset: 0;
  pointer-events: none;
  background-image:
    linear-gradient(rgba(39, 58, 138, 0.045) 1px, transparent 1px),
    linear-gradient(90deg, rgba(39, 58, 138, 0.045) 1px, transparent 1px),
    radial-gradient(circle at 20% 20%, rgba(6, 182, 212, 0.14), transparent 32%),
    radial-gradient(circle at 86% 8%, rgba(39, 58, 138, 0.12), transparent 28%);
  background-size: 48px 48px, 48px 48px, 100% 100%, 100% 100%;
}

.report-shell {
  position: relative;
  z-index: 1;
  max-width: 1240px;
  margin: 0 auto;
}

.hero-section {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 330px;
  gap: 28px;
  align-items: center;
  margin-bottom: 22px;
}

.hero-copy h1 {
  margin: 12px 0 0;
  color: var(--primary-strong);
  font-size: clamp(34px, 5vw, 58px);
  font-weight: 900;
  line-height: 1.02;
  letter-spacing: 0;
}

.hero-copy p {
  max-width: 780px;
  margin: 16px 0 0;
  color: #475569;
  font-size: 17px;
  line-height: 1.65;
}

.hero-badge,
.status-pill {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  width: fit-content;
  padding: 8px 12px;
  border: 1px solid rgba(37, 99, 235, 0.18);
  border-radius: 999px;
  background: rgba(255, 255, 255, 0.68);
  color: var(--primary);
  font-size: 12px;
  font-weight: 850;
  text-transform: uppercase;
}

.hero-visual {
  display: grid;
  gap: 12px;
  padding: 18px;
  border: 1px solid rgba(255, 255, 255, 0.72);
  border-radius: 8px;
  background: linear-gradient(145deg, rgba(255,255,255,0.86), rgba(240,249,255,0.68));
  box-shadow: var(--shadow);
  transform: perspective(900px) rotateY(-8deg) rotateX(4deg);
  animation: floatCard 6s ease-in-out infinite;
}

.hero-card-row {
  display: grid;
  grid-template-columns: 42px 1fr;
  gap: 12px;
  align-items: center;
  padding: 12px;
  border: 1px solid rgba(148, 163, 184, 0.22);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.82);
}

.hero-card-row svg {
  color: var(--accent);
}

.hero-card-row strong,
.file-panel-header strong,
.file-chip strong,
.nomination-chip strong {
  display: block;
  color: var(--ink);
  font-size: 14px;
}

.hero-card-row span,
.file-panel-header span,
.file-chip span,
.nomination-chip span {
  display: block;
  margin-top: 2px;
  color: var(--muted);
  font-size: 12px;
}

.stepper-card,
.dashboard-card,
.command-card {
  margin-bottom: 18px;
  border: 1px solid rgba(255, 255, 255, 0.72);
  border-radius: 8px;
  background: var(--glass);
  box-shadow: var(--shadow);
  backdrop-filter: blur(18px);
}

.stepper-card {
  display: grid;
  grid-template-columns: repeat(5, minmax(0, 1fr));
  gap: 8px;
  padding: 12px;
}

.step-item {
  display: flex;
  min-height: 54px;
  align-items: center;
  gap: 10px;
  padding: 10px;
  border: 1px solid transparent;
  border-radius: 8px;
  color: var(--muted);
  font-size: 13px;
  font-weight: 800;
  transition: 180ms ease;
}

.step-icon {
  display: grid;
  width: 32px;
  height: 32px;
  place-items: center;
  border-radius: 8px;
  background: rgba(148, 163, 184, 0.14);
}

.step-item.active,
.step-item.complete {
  color: var(--primary-strong);
  border-color: rgba(37, 99, 235, 0.18);
  background: rgba(255, 255, 255, 0.66);
}

.step-item.active .step-icon,
.step-item.complete .step-icon {
  background: linear-gradient(135deg, var(--primary), var(--accent));
  color: #ffffff;
}

.notice {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 16px;
  padding: 13px 15px;
  border-left: 4px solid;
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.86);
  box-shadow: 0 14px 30px rgba(15, 23, 42, 0.08);
  font-size: 14px;
  font-weight: 750;
}

.notice-error {
  border-color: var(--danger);
  color: #9f1239;
}

.notice-success {
  border-color: var(--success);
  color: #047857;
}

.dashboard-card {
  padding: 24px;
}

.section-heading {
  display: flex;
  justify-content: space-between;
  gap: 16px;
  margin-bottom: 18px;
}

.section-heading h2 {
  margin: 0;
  color: var(--ink);
  font-size: 20px;
  font-weight: 900;
}

.section-heading p {
  margin: 5px 0 0;
  color: var(--muted);
  font-size: 14px;
}

.source-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 18px;
  perspective: 1200px;
}

.source-card {
  display: grid;
  grid-template-columns: auto 1fr;
  gap: 16px;
  align-items: start;
  padding: 20px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.82);
  cursor: pointer;
  transform-style: preserve-3d;
  transition: transform 180ms ease, box-shadow 180ms ease, border-color 180ms ease;
}

.source-card:hover,
.source-card:focus-visible {
  transform: translateY(-5px) rotateX(2deg) rotateY(-2deg);
  box-shadow: 0 26px 50px rgba(37, 99, 235, 0.14);
  outline: none;
}

.source-card.selected {
  border-color: rgba(37, 99, 235, 0.72);
  box-shadow: 0 24px 56px rgba(37, 99, 235, 0.18), inset 0 0 0 1px rgba(6, 182, 212, 0.24);
}

.source-card.disabled {
  opacity: 0.58;
}

.source-icon {
  display: grid;
  width: 48px;
  height: 48px;
  place-items: center;
  border-radius: 8px;
  color: #ffffff;
}

.source-icon.teams {
  background: linear-gradient(135deg, #4f46e5, #2563eb);
}

.source-icon.webex {
  background: linear-gradient(135deg, #0891b2, #0f766e);
}

.source-card h3 {
  margin: 0;
  color: var(--ink);
  font-size: 19px;
}

.source-card p {
  margin: 6px 0 0;
  color: var(--muted);
  font-size: 13px;
  line-height: 1.45;
}

.file-type {
  display: inline-flex;
  margin-top: 10px;
  padding: 5px 8px;
  border-radius: 999px;
  background: rgba(6, 182, 212, 0.1);
  color: #0e7490;
  font-size: 11px;
  font-weight: 850;
}

.drop-zone,
.excel-zone {
  grid-column: 1 / -1;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  min-height: 138px;
  margin-top: 6px;
  padding: 20px;
  border: 1.5px dashed rgba(37, 99, 235, 0.35);
  border-radius: 8px;
  background: rgba(248, 250, 252, 0.72);
  color: #334155;
  text-align: center;
  cursor: pointer;
  transition: border-color 160ms ease, background 160ms ease;
}

.drop-zone:hover,
.excel-zone:hover {
  border-color: var(--accent);
  background: rgba(236, 254, 255, 0.62);
}

.drop-zone input,
.excel-zone input {
  max-width: 280px;
  margin: 10px 0;
}

.drop-zone svg,
.excel-zone svg {
  color: var(--primary);
}

.drop-zone strong,
.excel-zone strong {
  color: var(--ink);
  font-size: 15px;
}

.drop-zone span,
.excel-zone span {
  margin-top: 4px;
  color: var(--muted);
  font-size: 13px;
}

.file-panel {
  margin-top: 18px;
  padding: 16px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.62);
}

.file-panel-header {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 12px;
}

.file-chip-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px;
}

.file-chip,
.nomination-chip {
  display: grid;
  grid-template-columns: auto minmax(0, 1fr) auto;
  gap: 10px;
  align-items: center;
  padding: 10px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.82);
}

.file-chip svg,
.nomination-chip svg {
  color: var(--primary);
}

.file-chip strong,
.nomination-chip strong {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.file-chip button,
.nomination-chip button {
  display: grid;
  width: 30px;
  height: 30px;
  place-items: center;
  border: 0;
  border-radius: 8px;
  background: rgba(225, 29, 72, 0.08);
  color: var(--danger);
  cursor: pointer;
}

.options-grid {
  display: grid;
  grid-template-columns: minmax(240px, 0.8fr) minmax(340px, 1.2fr);
  gap: 20px;
  align-items: start;
}

.field-group,
.nomination-upload {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.nomination-upload {
  grid-column: 1 / -1;
}

.field-label {
  color: #334155;
  font-size: 13px;
  font-weight: 850;
}

.segmented,
.duration-pills {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.segmented button,
.duration-pills button {
  display: inline-flex;
  min-height: 42px;
  align-items: center;
  gap: 7px;
  padding: 9px 14px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.72);
  color: #475569;
  font-weight: 850;
  cursor: pointer;
  transition: 160ms ease;
}

.segmented button.selected,
.duration-pills button.selected {
  border-color: transparent;
  background: linear-gradient(135deg, var(--primary), var(--accent));
  color: #ffffff;
  box-shadow: 0 14px 28px rgba(37, 99, 235, 0.2);
}

.command-card {
  position: sticky;
  bottom: 14px;
  z-index: 2;
  padding: 16px;
}

.action-bar {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
}

.btn {
  position: relative;
  display: inline-flex;
  min-height: 46px;
  align-items: center;
  justify-content: center;
  gap: 9px;
  overflow: hidden;
  padding: 11px 18px;
  border: 1px solid transparent;
  border-radius: 8px;
  font-size: 14px;
  font-weight: 900;
  cursor: pointer;
  transition: transform 160ms ease, box-shadow 160ms ease, opacity 160ms ease;
}

.btn::after {
  content: "";
  position: absolute;
  top: 0;
  left: -120%;
  width: 80%;
  height: 100%;
  background: linear-gradient(90deg, transparent, rgba(255,255,255,0.34), transparent);
  transform: skewX(-18deg);
  transition: left 420ms ease;
}

.btn:not(:disabled):hover {
  transform: translateY(-2px);
}

.btn:not(:disabled):hover::after {
  left: 140%;
}

.btn:disabled {
  opacity: 0.48;
  cursor: not-allowed;
}

.btn-primary {
  background: linear-gradient(135deg, #273a8a, #2563eb, #06b6d4);
  color: #ffffff;
  box-shadow: 0 16px 30px rgba(37, 99, 235, 0.26);
}

.btn-secondary {
  background: linear-gradient(135deg, #059669, #10b981);
  color: #ffffff;
  box-shadow: 0 16px 30px rgba(16, 185, 129, 0.2);
}

.btn-danger-ghost {
  border-color: rgba(225, 29, 72, 0.22);
  background: rgba(255, 247, 247, 0.78);
  color: #be123c;
}

.spinner {
  width: 16px;
  height: 16px;
  border: 2px solid rgba(255, 255, 255, 0.45);
  border-top-color: #ffffff;
  border-radius: 999px;
  animation: spin 800ms linear infinite;
}

.summary-grid {
  display: grid;
  grid-template-columns: repeat(5, minmax(0, 1fr));
  gap: 14px;
  margin-bottom: 18px;
}

.summary-card {
  min-height: 144px;
  padding: 18px;
  border: 1px solid rgba(255,255,255,0.72);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.78);
  box-shadow: 0 18px 38px rgba(15, 23, 42, 0.09);
  animation: riseIn 420ms ease both;
  animation-delay: var(--delay);
}

.summary-card svg {
  color: var(--primary);
}

.summary-card strong {
  display: block;
  margin-top: 14px;
  color: var(--ink);
  font-size: 34px;
  line-height: 1;
}

.summary-card span {
  display: block;
  margin-top: 8px;
  color: #334155;
  font-size: 13px;
  font-weight: 900;
}

.summary-card small {
  display: block;
  margin-top: 5px;
  color: var(--muted);
  font-size: 12px;
}

.preview-table-wrap {
  max-height: 320px;
  overflow: auto;
  border: 1px solid var(--line);
  border-radius: 8px;
}

.preview-table {
  width: 100%;
  border-collapse: collapse;
  background: rgba(255,255,255,0.76);
}

.preview-table th {
  position: sticky;
  top: 0;
  z-index: 1;
  padding: 12px;
  background: #0f172a;
  color: #ffffff;
  font-size: 12px;
  text-align: left;
}

.preview-table td {
  padding: 12px;
  border-top: 1px solid rgba(148, 163, 184, 0.22);
  color: #334155;
  font-size: 13px;
}

.status-badge {
  display: inline-flex;
  padding: 5px 8px;
  border-radius: 999px;
  font-weight: 850;
}

.status-badge.present {
  background: rgba(16, 185, 129, 0.12);
  color: #047857;
}

.status-badge.absent {
  background: rgba(225, 29, 72, 0.1);
  color: #be123c;
}

.empty-state {
  display: grid;
  min-height: 170px;
  place-items: center;
  padding: 26px;
  border: 1px dashed var(--line);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.54);
  color: var(--muted);
  text-align: center;
}

.empty-state svg {
  color: var(--primary);
}

.empty-state strong {
  margin-top: 10px;
  color: var(--ink);
}

@keyframes gradientShift {
  0%, 100% { background-position: 0% 50%; }
  50% { background-position: 100% 50%; }
}

@keyframes floatCard {
  0%, 100% { transform: perspective(900px) rotateY(-8deg) rotateX(4deg) translateY(0); }
  50% { transform: perspective(900px) rotateY(-5deg) rotateX(5deg) translateY(-8px); }
}

@keyframes riseIn {
  from { opacity: 0; transform: translateY(14px); }
  to { opacity: 1; transform: translateY(0); }
}

@keyframes spin {
  to { transform: rotate(360deg); }
}

@media (max-width: 1080px) {
  .hero-section,
  .source-grid,
  .options-grid,
  .summary-grid {
    grid-template-columns: 1fr;
  }

  .hero-visual {
    max-width: 420px;
    transform: none;
  }

  .stepper-card {
    grid-template-columns: repeat(3, minmax(0, 1fr));
  }
}

@media (max-width: 720px) {
  .attendance-page {
    padding: 24px 14px 34px;
  }

  .hero-copy h1 {
    font-size: 34px;
  }

  .dashboard-card,
  .command-card {
    padding: 16px;
  }

  .file-chip-grid,
  .stepper-card {
    grid-template-columns: 1fr;
  }

  .section-heading,
  .file-panel-header {
    flex-direction: column;
  }

  .action-bar,
  .btn {
    width: 100%;
  }

  .command-card {
    position: static;
  }
}
</style>

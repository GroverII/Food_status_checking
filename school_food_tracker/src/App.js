import React, { useState, useEffect } from 'react';
import { Modal, Button } from 'react-bootstrap';
import { format, startOfYear, endOfYear } from 'date-fns';
import DatePicker from 'react-datepicker';
import 'react-datepicker/dist/react-datepicker.css';
import { getAllClasses, getStudentData, findStudent } from './api.js';
import './styles.css';

const SchoolFoodComponent = () => {
    const [startDate, setStartDate] = useState(null);
    const [endDate, setEndDate] = useState(null);
    const [classValue, setClassValue] = useState('');
    const [nameValue, setNameValue] = useState('');
    const [surnameValue, setSurnameValue] = useState('');
    const [codeValue, setCodeValue] = useState('');
    const [hasContract, setHasContract] = useState(null);
    const [isPaid, setIsPaid] = useState(null);
    const [classes, setClasses] = useState([]);
    const [students, setStudents] = useState([]);
    const [showModal, setShowModal] = useState(false);


    useEffect(() => {
        try {
            getAllClasses()
                .then(data => setClasses(data))
                .catch(error => {
                    console.error('Error fetching classes:', error);
                    // Дополнительная обработка ошибки, если необходимо
                });
        } catch (error) {
            console.error('Error in useEffect:', error);
            // Обработка ошибки при выполнении useEffect
        }
    }, []);


    const handleExcelData = async () => {
        try {
            if (startDate && endDate && startDate > endDate) {
                alert('Start Month must be greater than or equal to End Month.');
                return;
            }

            const requestBody = {
                StartDate: startDate ? format(startDate, 'yyyy-MM-dd') : format(startOfYear(new Date()), 'yyyy-MM-dd'),
                EndDate: endDate ? format(endDate, 'yyyy-MM-dd') : format(endOfYear(new Date()), 'yyyy-MM-dd'),
                Class: classValue,
                Name: nameValue,
                Surname: surnameValue,
                Code: codeValue,
                HasContract: hasContract,
                IsPaid: isPaid,
            };

            const excelBlob = await getStudentData(requestBody);

            const excelUrl = URL.createObjectURL(excelBlob);

            const link = document.createElement('a');
            link.href = excelUrl;
            link.download = 'StudentData.xlsx';
            link.style.display = 'none';

            document.body.appendChild(link);
            link.click();

            document.body.removeChild(link);
            URL.revokeObjectURL(excelUrl);

            alert('Data successfully downloaded');
        } catch (error) {
            console.error('Error downloading data:', error.message);
            alert('Error downloading data. Please check the console for more details.');
        }
    };


    const handleAtzimetSkolenu = async () => {
        try {
            if (!codeValue) throw new Error('Nav kodu!');
            const result = await findStudent(codeValue);

            console.log('Result:', result); // Log the result object

            // Предполагаем, что результат в формате JSON
            const { status, personInfo } = JSON.parse(result) || {};

            if (personInfo) {
                alert(`Status: ${status}\nName: ${personInfo.name || 'N/A'}\nSurname: ${personInfo.surname || 'N/A'}\nClass: ${personInfo.class || 'N/A'}\nHas Contract: ${personInfo.hasContract}\nIs Paid: ${personInfo.isPaid}`);

                const currentTime = formatTime();

                // Обновляем состояние студентов, добавляя нового студента в массив
                setStudents(prevStudents => [
                    ...prevStudents,
                    {
                        name: personInfo.name || 'N/A',
                        surname: personInfo.surname || 'N/A',
                        class: personInfo.class || 'N/A',
                        hasContract: personInfo.hasContract,
                        isPaid: personInfo.isPaid,
                        timestamp: currentTime
                    }
                ]);
            } else {
                throw new Error('PersonInfo is undefined or null');
            }
        } catch (error) {
            console.error('Error finding data:', error.message);
            alert('Error finding data(' + error.message + ')');
        }
    };

    const formatTime = () => {
        const now = new Date();
        const hours = now.getHours().toString().padStart(2, '0');
        const minutes = now.getMinutes().toString().padStart(2, '0');
        const seconds = now.getSeconds().toString().padStart(2, '0');
        return `${hours}:${minutes}:${seconds}`;
    };

    const handleClearList = () => {
        // Очищаем список студентов
        setStudents([]);
    };


    const isWeekday = (date) => {
        const day = date.getDay();
        return day !== 0; // Фильтруем воскресенье (0)
    };

    const getCurrentDate = () => {
        const today = new Date();
        const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        return today.toLocaleDateString('en-US', options);
    };

    function calculateWidth(str) {
        return [...str].reduce((width, char) => {
            // Для широких символов с диакритическими знаками
            if (char.match(/[āēīūčšģķļņž]/i)) {
                return width + 2; // Предполагаем, что эти символы в два раза шире
            }
            return width + 1;
        }, 0);
    }

    const exportToTxt = () => {
        // Функция для выравнивания текста в столбцах
        const alignText = (text, maxLength) => {
            const spacesToAdd = maxLength - text.length;
            const padding = ' '.repeat(spacesToAdd);
            return `${text}${padding}`;
        };

        // Найти максимальную длину для каждого столбца
        const maxLengths = {
            index: String(students.length).length + 1,
            name: Math.max(...students.map(student => calculateWidth(student.name))),
            surname: Math.max(...students.map(student => calculateWidth(student.surname))),
            class: Math.max(...students.map(student => calculateWidth(student.class))),
            hasContract: calculateWidth('Līgums'),
            isPaid: calculateWidth('Samaksāts'),
            timestamp: Math.max(...students.map(student => calculateWidth(student.timestamp)))
        };

        // Создать отформатированный текст
        const txtContent = students.map((student, index) => (
            `${alignText(student.timestamp, maxLengths.timestamp)}\t${alignText(index + 1 + '.', maxLengths.index)} \t${alignText(student.name, maxLengths.name)} \t${alignText(student.surname, maxLengths.surname)} \t${alignText(student.class, maxLengths.class)} \tLīgums ${alignText(student.hasContract ? 'ir' : 'nav', maxLengths.hasContract)} \tSamaksāts ${alignText(student.isPaid ? 'ir' : 'nav', maxLengths.isPaid)}\n`
        )).join('');

        // Создать и скачать файл
        const blob = new Blob([txtContent], { type: 'text/plain' });
        const url = URL.createObjectURL(blob);

        const link = document.createElement('a');
        link.href = url;
        link.download = 'StudentData.txt';
        link.style.display = 'none';

        document.body.appendChild(link);
        link.click();

        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    };




    return (
        <div className={`slide ${showModal ? 'active blur' : ''}`}>
            <div className="overlay" onClick={() => setShowModal(false)}></div>
            <div className="div">
                <button className="overlap" onClick={handleAtzimetSkolenu}>
                    Atzīmējiet studentu
                </button>
                <div className="text-wrapper-9">Kods:</div>
                <div className="overlap-2">
                    <input
                        className="v-rds-input"
                        type="number"
                        value={codeValue}
                        onChange={(e) => setCodeValue(e.target.value)}
                        onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                                e.preventDefault();
                                handleAtzimetSkolenu();
                            }
                        }}
                    />
                </div>

                <div style={{ textAlign: 'center', margin: '20px 0' }}>
                    <h2>{getCurrentDate()}</h2>
                </div>
                <table className="student-table">
                    <thead>
                        <tr>
                            <th>Nr.</th>
                            <th>Vārds</th>
                            <th>Uzvārds</th>
                            <th>Klase</th>
                            <th>Līgums</th>
                            <th>Samaksāts</th>
                            <th>Laiks</th>
                        </tr>
                    </thead>
                    <tbody>
                        {/* Проходим по вашим данным и создаем строки таблицы */}
                        {students.map((student, index) => (
                            <tr key={index}>
                                <td>{index + 1}</td>
                                <td>{student.name}</td>
                                <td>{student.surname}</td>
                                <td>{student.class}</td>
                                <td>{student.hasContract ? 'Ir' : 'Nav'}</td>
                                <td>{student.isPaid ? 'Ir' : 'Nav'}</td>
                                <td>{student.timestamp}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>
                <button onClick={handleClearList}>Очистить список</button>

                <button onClick={exportToTxt}>Экспортировать в TXT</button>

                <button onClick={() => setShowModal(true)}>Открыть модальное окно</button>

                <Modal show={showModal} onHide={() => setShowModal(false)} backdrop="static" keyboard={false}>
                    <Modal.Body>
                        <div className="text-wrapper-3">Datums:</div>
                        <div className="text-wrapper-4">Klases:</div>
                        <div className="text-wrapper-5">Ir samaksāts:</div>
                        <div className="text-wrapper-6">Līgums:</div>
                        <div className="text-wrapper-7">Vārds:</div>
                        <div className="text-wrapper-8">Uzvārds:</div>
                        <div className="v-rds-wrapper">
                            <input
                                className="v-rds-input"
                                type="text"
                                value={surnameValue}
                                onChange={(e) => setSurnameValue(e.target.value)}
                            />
                        </div>

                        <div className="v-rds-wrapper">
                            <input
                                className="v-rds-input"
                                type="text"
                                value={surnameValue}
                                onChange={(e) => setSurnameValue(e.target.value)}
                            />
                        </div>
                        <div className="overlap-3">
                            <input
                                className="v-rds-input"
                                type="text"
                                value={nameValue}
                                onChange={(e) => setNameValue(e.target.value)}
                            />
                        </div>
                        <select
                            className="div-wrapper"
                            onChange={(e) => setIsPaid(e.target.value === 'null' ? null : e.target.value === 'true')}
                            value={isPaid === null ? 'null' : isPaid ? 'true' : 'false'}
                        >
                            <option value="null">Visi</option>
                            <option value="true">Ir</option>
                            <option value="false">Nav</option>
                        </select>
                        <select
                            className="overlap-4"
                            onChange={(e) => setClassValue(e.target.value)}
                            value={classValue}
                        >
                            <option value="">Visi</option>
                            {classes.map((classItem) => (
                                <option key={classItem} value={classItem}>
                                    {classItem}
                                </option>
                            ))}
                        </select>

                        <select
                            className="overlap-5"
                            onChange={(e) => setHasContract(e.target.value === 'null' ? null : e.target.value === 'true')}
                            value={hasContract === null ? 'null' : hasContract ? 'true' : 'false'}
                        >
                            <option value="null">Visi</option>
                            <option value="true">Ir</option>
                            <option value="false">Nav</option>
                        </select>
                        <div className="overlap-6">
                            <DatePicker
                                selected={endDate}
                                onChange={(date) => setEndDate(date)}
                                dateFormat="dd/MM/yyyy"
                                className="date-picker"
                                minDate={startDate}
                                filterDate={isWeekday} // Используем фильтр для исключения выходных дней
                            />
                        </div>

                        <div className="overlap-7">
                            <DatePicker
                                selected={startDate}
                                onChange={(date) => setStartDate(date)}
                                dateFormat="dd/MM/yyyy"
                                className="date-picker"
                                maxDate={endDate}
                                filterDate={isWeekday} // Используем фильтр для исключения выходных дней
                            />
                        </div>
                    </Modal.Body>
                    <Modal.Footer>
                        <Button variant="secondary" onClick={() => setShowModal(false)}>
                            Закрыть
                        </Button>
                        <Button className="export-excel" onClick={handleExcelData}>
                            Importēt uz Excel
                        </Button>
                    </Modal.Footer>
                </Modal>
            </div>
        </div>
    );
};

export default SchoolFoodComponent;

const XLSX = require('../../xlsx');
const axios = require('axios');

// const wb = XLSX.readFile('./hertz.xlsm', {
// 	cellStyles: true,
// 	cellDates: true,
// 	bookVBA: true,
// 	cellFormula: true
// });

async function process() {
	const response = await axios.get(
		'https://terraceag-file-upload.s3.us-east-2.amazonaws.com/34e84225-c365-405a-a371-ec50f814ea45-hertz.xlsm',
		{ responseType: 'arraybuffer' }
	);

	const wb = XLSX.read(response.data, {
		cellStyles: true,
		cellDates: true,
		bookVBA: true,
		cellFormula: true
	});

	const sale = require('./sales.json')[0];

	const ws = wb.Sheets[wb.SheetNames[0]];

	ws['B4'].v = sale.effective_sale_date ? sale.effective_sale_date : '';
	ws['B5'].v = sale.recording_date ? sale.recording_date : '';
	ws['B6'].v = sale.sale_price ? sale.sale_price : 0;
	ws['B7'].v = sale.adjusted_sale_price ? sale.adjusted_sale_price : '';
	ws['B8'].v = sale.total_acres ? sale.total_acres : 0;
	ws['B9'].v = sale.price_per_acre ? sale.price_per_acre : 0;
	ws['B10'].v = sale.cropland_csr2 ? sale.cropland_csr2 : 0;

	ws['D4'].v = sale.county ? sale.county : '';
	ws['D5'].v = sale.township ? sale.township : '';
	ws['D6'].v = sale.grantee ? sale.grantee : '';
	ws['D7'].v = sale.grantor ? sale.grantor : '';
	ws['D8'].v = sale.instrument_number ? sale.instrument_number : '';
	ws['D9'].v = sale.abbreviated_legal_description
		? `Abbreviated Legal Description: ${sale.abbreviated_legal_description}`
		: '';

	ws['E13'].v = sale.data_source ? sale.data_source : '';

	ws['E17'].v = sale.soil && sale.soil[0] ? sale.soil[0].acres : 0;
	ws['E18'].v = sale.soil && sale.soil[1] ? sale.soil[1].acres : 0;
	ws['E19'].v = sale.soil && sale.soil[2] ? sale.soil[2].acres : 0;
	ws['E20'].v = sale.soil && sale.soil[3] ? sale.soil[3].acres : 0;
	ws['E21'].v = sale.soil && sale.soil[4] ? sale.soil[4].acres : 0;
	ws['E22'].v = sale.soil && sale.soil[5] ? sale.soil[5].acres : 0;
	ws['E23'].v = sale.soil && sale.soil[5] ? sale.soil[5].acres : 0;
	ws['E24'].v = sale.soil && sale.soil[5] ? sale.soil[5].acres : 0;

	ws['F17'].v = sale.soil && sale.soil[0] ? sale.soil[0].dollar_acre : 0;
	ws['F18'].v = sale.soil && sale.soil[1] ? sale.soil[1].dollar_acre : 0;
	ws['F19'].v = sale.soil && sale.soil[2] ? sale.soil[2].dollar_acre : 0;
	ws['F20'].v = sale.soil && sale.soil[3] ? sale.soil[3].dollar_acre : 0;
	ws['F21'].v = sale.soil && sale.soil[4] ? sale.soil[4].dollar_acre : 0;
	ws['F22'].v = sale.soil && sale.soil[5] ? sale.soil[5].dollar_acre : 0;
	ws['F23'].v = sale.soil && sale.soil[5] ? sale.soil[5].dollar_acre : 0;
	ws['F24'].v = sale.soil && sale.soil[5] ? sale.soil[5].dollar_acre : 0;

	ws['H17'].v = sale.soil && sale.soil[0] ? sale.soil[0].total : 0;
	ws['H18'].v = sale.soil && sale.soil[1] ? sale.soil[1].total : 0;
	ws['H19'].v = sale.soil && sale.soil[2] ? sale.soil[2].total : 0;
	ws['H20'].v = sale.soil && sale.soil[3] ? sale.soil[3].total : 0;
	ws['H21'].v = sale.soil && sale.soil[4] ? sale.soil[4].total : 0;
	ws['H22'].v = sale.soil && sale.soil[5] ? sale.soil[5].total : 0;
	ws['H23'].v = sale.soil && sale.soil[5] ? sale.soil[5].total : 0;
	ws['H24'].v = sale.soil && sale.soil[5] ? sale.soil[5].total : 0;

	ws['D28'].v = sale.improvements && sale.improvements[0] ? sale.improvements[0].improvements : '';
	ws['D29'].v = sale.improvements && sale.improvements[1] ? sale.improvements[1].improvements : '';
	ws['D30'].v = sale.improvements && sale.improvements[2] ? sale.improvements[2].improvements : '';

	ws['E28'].v = sale.improvements && sale.improvements[0] ? sale.improvements[0].size : 0;
	ws['E29'].v = sale.improvements && sale.improvements[1] ? sale.improvements[1].size : 0;
	ws['E30'].v = sale.improvements && sale.improvements[2] ? sale.improvements[2].size : 0;

	ws['F28'].v = sale.improvements && sale.improvements[0] ? sale.improvements[0].unit : '';
	ws['F29'].v = sale.improvements && sale.improvements[1] ? sale.improvements[1].unit : '';
	ws['F30'].v = sale.improvements && sale.improvements[2] ? sale.improvements[2].unit : '';

	ws['G28'].v = sale.improvements && sale.improvements[0] ? sale.improvements[0].dollar_unit : 0;
	ws['G29'].v = sale.improvements && sale.improvements[1] ? sale.improvements[1].dollar_unit : 0;
	ws['G30'].v = sale.improvements && sale.improvements[2] ? sale.improvements[2].dollar_unit : 0;

	ws['B36'].v = sale.topography ? sale.topography : '';
	ws['B37'].v = sale.drainage ? sale.drainage : '';
	ws['B38'].v = sale.farming_efficiency ? sale.farming_efficiency : '';
	ws['B39'].v = sale.other_comments ? sale.other_comments : '';

	ws['C45'].v = sale.income && sale.income[0] ? sale.income[0].dollar_unit : 0;
	ws['C46'].v = sale.income && sale.income[1] ? sale.income[1].dollar_unit : 0;
	ws['C47'].v = sale.income && sale.income[2] ? sale.income[2].dollar_unit : 0;
	ws['C48'].v = sale.income && sale.income[3] ? sale.income[3].dollar_unit : 0;
	ws['C49'].v = sale.income && sale.income[4] ? sale.income[4].dollar_unit : 0;
	ws['C50'].v = sale.income && sale.income[5] ? sale.income[5].dollar_unit : 0;

	ws['G45'].v = sale.taxes ? sale.taxes : 0;

	ws['D55'].v = sale.latitude ? sale.latitude : '';
	ws['D56'].v = sale.longitude ? sale.longitude : '';

	XLSX.writeFile(wb, 'generated.xlsm', {
		bookType: 'xlsm'
		// cellStyles: true,
		// cellDates: true,
		// bookVBA: true,
		// cellFormula: true
	});
}

process();

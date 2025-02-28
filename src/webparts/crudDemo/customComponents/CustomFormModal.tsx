import * as React from 'react';
import './CustomFormModal.css';
import IEmployeeDetails from '../../../models/IEmplyeeDetails';

interface ModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSave: (formData: IEmployeeDetails) => void;
  onUpdate: (formData: IEmployeeDetails) => void;
  item: IEmployeeDetails;
  isUpdate: boolean;
}

const CustomFormModal: React.FC<ModalProps> = ({
  isOpen,
  onClose,
  onSave,
  onUpdate,
  item,
  isUpdate,
}) => {
  const [isModalOpen, setIsModalOpen] = React.useState(isOpen);

  const [formData, setFormData] = React.useState<IEmployeeDetails>(item);

  const [dataList, setDataList] = React.useState<IEmployeeDetails[]>([]);

  // get initial state after initial render
  React.useEffect(() => {
    setIsModalOpen(isOpen);
    setFormData(item);
  }, [isOpen]);

  if (!isModalOpen) return null;

  const handleInputChange = (
    e: React.ChangeEvent<
      HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement
    >
  ) => {
    let { name, value } = e.target;

    //input validation for numbers only
    if (name === 'Salary') {
      value = value
        .replace(/[^0-9.]/g, '')
        .replace(/(\..*?)\..*/g, '$1')
        .replace(/^0[^.]/, '0');
    }
    setFormData({ ...formData, [name]: value });
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (
      formData.FirstName &&
      formData.LastName &&
      formData.Gender &&
      formData.Salary
    ) {
      setDataList([...dataList, formData]);
      //Send form data to parent component
      if (isUpdate) {
        onUpdate(formData);
      } else {
        onSave(formData);
      }

      //Reset form
      setFormData({
        Id: 0,
        FirstName: '',
        LastName: '',
        Gender: '',
        Salary: '',
      });
      //Close modal
      onClose();
    }
  };

  return (
    <>
      {isModalOpen && (
        <div className="modal-overlay">
          <div className="modal">
            <h2>Enter Details</h2>
            <form onSubmit={handleSubmit}>
              <input
                type="text"
                name="FirstName"
                placeholder="First Name"
                value={formData.FirstName}
                onChange={handleInputChange}
                required
              />
              <input
                type="text"
                name="LastName"
                placeholder="Last Name"
                value={formData.LastName}
                onChange={handleInputChange}
                required
              />
              <select
                name="Gender"
                value={formData.Gender}
                onChange={handleInputChange}
                required
              >
                <option value="" disabled selected hidden>
                  Gender
                </option>
                <option value="Male">Male</option>
                <option value="Female">Female</option>
                <option value="Neutral">Neutral</option>
              </select>
              <input
                type="text"
                name="Salary"
                placeholder="Salary"
                value={formData.Salary}
                onChange={handleInputChange}
                required
              ></input>
              <div className="modal-buttons">
                <button type="button" className="close-btn" onClick={onClose}>
                  Cancel
                </button>
                <button type="submit" className="submit-btn">
                  {isUpdate ? 'Update' : 'Submit'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
};

export default CustomFormModal;

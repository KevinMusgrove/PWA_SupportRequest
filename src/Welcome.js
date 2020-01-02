import React from 'react';
import { Button } from 'reactstrap';
import { sendEmail } from './GraphService';
import config from './Config';
import { FilePond, registerPlugin } from 'react-filepond';
import 'filepond/dist/filepond.min.css';
import FilePondPluginFileEncode from 'filepond-plugin-file-encode';
import FilePondPluginImagePreview from 'filepond-plugin-image-preview';
import 'filepond-plugin-image-preview/dist/filepond-plugin-image-preview.min.css';
import notifier from "simple-react-notifications";
import "simple-react-notifications/dist/index.css";
import "react-loader-spinner/dist/loader/css/react-spinner-loader.css"
import Loader from 'react-loader-spinner'
import { Select } from 'antd';

let fileUpload = React.createRef();
registerPlugin(FilePondPluginFileEncode);
registerPlugin(FilePondPluginImagePreview);
const { Option } = Select;

class RequestForm extends React.Component{
  constructor(props) {
    super(props);
    this.submitFormMethod = this.submitFormMethod.bind(this);
    this.ClearForm = this.ClearForm.bind(this);    
    this.selectDevice = this.selectDevice.bind(this);
    this.selectPriority = this.selectPriority.bind(this);
    this.selectLocation = this.selectLocation.bind(this);
    this.selectPhone = this.selectPhone.bind(this);
    this.state = {
      isLoading: false,
      deviceValue:"",
      priorityValue:"Medium",
      locationValue:"",
      phoneValue: ""
    };
  }

  devices = ["Laptop","Desktop","Monitor","Bluetooth headset","Wired headset","Cell phone","Desk phone","Keyboard","Mouse","Speakers","Printer","Tablet","Not listed"];
  locations = ["Location 1"," Location 2"," Location 3"];
  priorites = ["High","Medium","Low"]

  async submitFormMethod(event,props){    
    try{                          
      event.preventDefault();            
      var accessToken = await props.UserAgentApplication.acquireTokenSilent({
        scopes: config.scopes
      });
      if (accessToken) {
        this.setState({isLoading:true}); 
        let firstName = document.getElementById("tbFirstname").value;        
        let lastName = document.getElementById("tbLastname").value;
        let email = document.getElementById("tbEmail").value;
        let issueDesc = document.getElementById("tbIssueDesc").value;        
        let location = this.state.locationValue;        
        let phone = this.state.phoneValue || this.props.user.cellPhone || this.props.user.bussinessPhone || document.getElementById("tbPhone").value;
        let device = this.state.deviceValue;
        let priority = this.state.priorityValue;        
        let filePondObjects = fileUpload.current.getFiles();
        let files = [];
        if(firstName.trim() == "" || lastName.trim() == "" || location.trim() == "" || email.trim() == "" || phone.trim() == "" || device.trim() == "" || priority.trim() == "" || issueDesc.trim() == ""){
          notifier.error(`Error: Please fill out all required fields.`);
          this.setState({isLoading:false}); 
          return;
        }
        filePondObjects.forEach(file => {                  
            files.push({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": `${file.filename}`,
                "contentBytes": `${file.getFileEncodeBase64String()}`
            })
        });
        await sendEmail(accessToken,firstName,lastName,location,email,phone,device,priority,issueDesc,props.managerEmail,files);
        //Show success method.
        notifier.configure({
          autoClose: 3000,
          width: 275,
          position: "top-right",
          delay: 0,
          closeOnClick: true,
          pauseOnHover: true,
          onlyLast: false,
          rtl: false,
          newestOnTop: true,
          animation: {
            in: "zoomIn",
            out: "zoomOut",
            duration: 400
          }
        });
        notifier.success("Request successfully sent.");  
        //notifier.info("Someone from ATB will be reaching out as soon as posible.")       
        this.ClearForm();
      }
      else{
        // have user reauthenticate
        props.authButtonMethod();
      }
    }
    catch(err){   
      debugger;   
        //Show error  
        notifier.configure({
          autoClose: 3000,
          width: 275,
          position: "top-right",
          delay: 0,
          closeOnClick: true,
          pauseOnHover: true,
          onlyLast: false,
          rtl: false,
          newestOnTop: true,
          animation: {
            in: "zoomIn",
            out: "zoomOut",
            duration: 400
          }
        });      
        notifier.error(`Error:  ${err.body}`);
        this.setState({isLoading:false}); 
    }
}

 ClearForm(){    
  var tbPhone = document.getElementById("tbPhone");
  if(tbPhone)
  {
    tbPhone.value = "";
  }
  this.setState({
    isLoading: false,deviceValue:"",priorityValue:"Medium",locationValue:"",phoneValue:""
  });
  document.getElementById("tbIssueDesc").value = "";    
  if(fileUpload.current.getFiles().length != 0){                
      fileUpload.current.getFiles().forEach(file => {
          fileUpload.current.removeFile(file);
      });        
  }
}

selectDevice = value => {
   this.setState((prevState, props) =>{
     return{isLoading:prevState.isLoading,priorityValue:prevState.priorityValue,locationValue:prevState.locationValue,phoneValue:prevState.phoneValue,deviceValue:value}
   })
}

selectLocation = value => {
  this.setState((prevState, props) =>{
    return{isLoading:prevState.isLoading,priorityValue:prevState.priorityValue,deviceValue:prevState.deviceValue,phoneValue:prevState.phoneValue,locationValue:value}
  })
}

selectPriority = value => {
  this.setState((prevState, props) =>{
    return{isLoading:prevState.isLoading,deviceValue:prevState.deviceValue,locationValue:prevState.locationValue,phoneValue:prevState.phoneValue,priorityValue:value}
  })
}
selectPhone= value => {
  this.setState((prevState, props) =>{
    return{isLoading:prevState.isLoading,deviceValue:prevState.deviceValue,locationValue:prevState.locationValue,priorityValue:prevState.priorityValue,phoneValue:value}
  })
}

  render(){
  console.log(this.props.user)  
  if (this.props.isAuthenticated) {
    return (
      <div>
        <h4 className="display-10 text-center mb-3" id="appHeader">
            New I.T. Support Request
        </h4>
      <div className="container">
        <div className="row">
            <div className="col-md-6 offset-md-3">
              <form>
                <div className="form-group">    
                    <label className="inputLabel">* First Name</label>   
                    <input type="text" className="form-control" id="tbFirstname" defaultValue={this.props.user.firstName} placeholder="First Name" required={true}/>                                                            
                </div>
                <div className="form-group">
                  <label className="inputLabel">* Last Name</label>   
                  <input type="text" className="form-control" id="tbLastname" defaultValue={this.props.user.lastName} placeholder="Last Name" required/>
                </div>
                    <div className="form-group">
                        <label className="inputLabel">* Email</label>   
                        <input type="text" className="form-control" id="tbEmail" defaultValue={this.props.user.email} placeholder="Email" required/>
                    </div>
                    <div className="form-group">
                        <label className="inputLabel">* Phone</label>   
                        {this.props.user.cellPhone && this.props.user.bussinessPhone ? 
                            <Select                      
                              value={this.state.phoneValue || this.props.user.cellPhone || this.props.user.bussinessPhone}
                              defaultValue={this.props.user.cellPhone || this.props.user.bussinessPhone}  
                              showSearch={true}
                              size="large"
                              placeholder="Select Phone Number"                                            
                              filterOption={true}                      
                              onChange={this.selectPhone}
                              style={{ width: '100%' }}
                            >                            
                              <Option key={this.props.user.cellPhone}>{this.props.user.cellPhone}</Option>
                              <Option key={this.props.user.bussinessPhone}>{this.props.user.bussinessPhone}</Option>                            
                            </Select>
                            :<input type="text" className="form-control" id="tbPhone" defaultValue={this.props.user.cellPhone || this.props.user.bussinessPhone} placeholder="Phone Number" required/>
                        }                        
                    </div>
                    <div className="form-group">
                      <label className="inputLabel">* Location</label>   
                      <Select                      
                        value={this.state.locationValue}
                        showSearch={true}
                        size="large"
                        placeholder="Select Location"                                            
                        filterOption={true}                      
                        onChange={this.selectLocation}
                        style={{ width: '100%' }}
                      >
                        {this.locations.map(d => (
                          <Option key={d}>{d}</Option>
                        ))}
                      </Select>
                    </div>                                  
                    <div className="form-group">
                      <label className="inputLabel">* Issue Priority</label>   
                      <Select                      
                        value={this.state.priorityValue}
                        defaultValue="Medium"
                        showSearch={true}
                        size="large"
                        placeholder="Select Priority Level"                                            
                        filterOption={true}                      
                        onChange={this.selectPriority}
                        style={{ width: '100%' }}
                      >
                        {this.priorites.map(d => (
                          <Option key={d}>{d}</Option>
                        ))}
                      </Select>
                    </div>
                    <div className="form-group">
                      <label className="inputLabel">* Device with Issue</label>   
                      <Select                      
                        value={this.state.deviceValue}
                        showSearch={true}
                        size="large"
                        placeholder="Select Device"                                            
                        filterOption={true}                      
                        onChange={this.selectDevice}
                        style={{ width: '100%' }}
                      >
                        {this.devices.map(d => (
                          <Option key={d}>{d}</Option>
                        ))}
                      </Select>
                    </div>
                    <div className="form-group">
                        <label className="inputLabel">* Issue Description</label>   
                        <textarea className="form-control" rows="3" id="tbIssueDesc" placeholder="Issue Description" required></textarea> 
                    </div>
                    <label className="inputLabel">Attachments</label>   
                    <FilePond id="fileControl" allowMultiple={true} ref={fileUpload}/>
                    <Loader
                      type="TailSpin"
                      color="lightGray"
                      height={50}
                      width={50}
                      className="loader"                      
                      visible={this.state.isLoading}   
                    />
                    <div className="text-center mb-3" id="btnSection">
                        <input type="button" value="Reset Form" id="btnClear" onClick={() => this.ClearForm()} className="btn btn-danger"/>
                        <input type="submit" value="Submit Form" id="btnSubmit" onClick={(e) => this.submitFormMethod(e,this.props)} className="btn btn-primary"/>
                    </div> 
              </form>
            </div>
        </div>
      </div>
      </div>
    );
  }
  // Not authenticated, present a sign in button
  return (
    <div id="signIn"> 
        <p>Welcome, please sign in below to create a new support request.</p>     
        <Button id="btnSignInMain" onClick={this.props.authButtonMethod}>Sign in with Microsoft</Button>
    </div>  
  );
}
}
export default class Welcome extends React.Component {
  render() {
    return (
        <div>
          <RequestForm
          isAuthenticated={this.props.isAuthenticated}
          user={this.props.user}          
          managerEmail={this.props.managerEmail}
          UserAgentApplication={this.props.UserAgentApplication}
          authButtonMethod={this.props.authButtonMethod} />          
        </div>
    );    
  }
}